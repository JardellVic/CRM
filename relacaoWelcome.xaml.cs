using ClosedXML.Excel;
using OfficeOpenXml;
using CRM.conexao;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Windows;
using System.Text.RegularExpressions;

namespace CRM
{
    public partial class relacaoWelcome : Window
    {
        private DataTable _dataTable;
        private conexaoMouraWelcome dbHelper;

        public relacaoWelcome()
        {
            InitializeComponent();
            SetupDates();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            dbHelper = new conexaoMouraWelcome();
            _dataTable = new DataTable();
        }

        private void SetupDates()
        {
            DateTime today = DateTime.Now;
            DateTime startDate = today.AddDays(-1); // Um dia antes
            txtDataInicial.Text = startDate.ToString("dd/MM/yyyy");
        }

        private async void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DateTime startDate = DateTime.ParseExact(txtDataInicial.Text, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture).Date;
                DateTime endDate = startDate.AddDays(1).AddSeconds(-1);

                progressBar.Visibility = Visibility.Visible;
                progressBar.IsIndeterminate = true;

                _dataTable = await Task.Run(() => dbHelper.FetchData(startDate, endDate));

                if (_dataTable != null && _dataTable.Rows.Count > 0)
                {
                    lblTotalRecords.Content = $"Total de registros: {_dataTable.Rows.Count}";
                    btnExportarExcel.IsEnabled = true;
                }
                else
                {
                    lblTotalRecords.Content = "Nenhum registro encontrado.";
                    btnExportarExcel.IsEnabled = false;
                }

                progressBar.Visibility = Visibility.Collapsed;
            }
            catch (FormatException ex)
            {
                MessageBox.Show($"Erro ao converter a data: {ex.Message}");
            }
        }

        private async void btnExportarExcel_Click(object sender, RoutedEventArgs e)
        {
            if (_dataTable == null || _dataTable.Rows.Count == 0)
            {
                MessageBox.Show("Nenhum dado para exportar.");
                return;
            }
           
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string relacoesPath = Path.Combine(desktopPath, "Relações");
            if (!Directory.Exists(relacoesPath))
            {
                Directory.CreateDirectory(relacoesPath);
            }
            string outputPath = Path.Combine(relacoesPath, "Welcome.xlsx");
            string? directoryPath = Path.GetDirectoryName(outputPath);



            if (directoryPath != null)
            {
                Directory.CreateDirectory(directoryPath);
            }
            else
            {
                MessageBox.Show("Não foi possível determinar o caminho do diretório para salvar o arquivo.", "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            DataTable filteredTable = FilterColumns(_dataTable);

            progressBar.Visibility = Visibility.Visible;
            progressBar.IsIndeterminate = true;

            await Task.Run(() => SaveToExcel(filteredTable, outputPath));

            progressBar.Visibility = Visibility.Collapsed;

            MessageBox.Show($"Arquivo salvo em {outputPath}!", "Concluído", MessageBoxButton.OK);
            this.Close();
        }

        /*  private DataTable FilterColumns(DataTable dataTable)
         {
             // Retorna o DataTable original sem filtrar as colunas
             return dataTable;
         } */

        private DataTable FilterColumns(DataTable dataTable)
        {
            DataTable filteredTable = new DataTable();

            filteredTable.Columns.Add("nome", typeof(string));
            filteredTable.Columns.Add("Data_Cadastro", typeof(DateTime));
            filteredTable.Columns.Add("fone", typeof(string));
            filteredTable.Columns.Add("fone2", typeof(string));

            var filteredRows = dataTable.AsEnumerable()
                .Where(row => !row["Nome"].ToString().Contains("#") &&
                              !row["Nome"].ToString().Contains("@") &&
                              !row["Nome"].ToString().Contains("&") &&
                              !row["Nome"].ToString().Contains("MERCADO LIVRE") &&
                              !row["Nome"].ToString().Contains("CONSUMIDOR FINAL"));


            foreach (var row in filteredRows)
            {
                DataRow newRow = filteredTable.NewRow();
                newRow["Nome"] = row["Nome"];
                newRow["Data_Cadastro"] = row["Data_Cadastro"];
                newRow["Fone"] = FormatPhoneNumber(row["Fone"].ToString());
                newRow["Fone2"] = FormatPhoneNumber(row["Fone2"].ToString());
                filteredTable.Rows.Add(newRow);
            }

            return filteredTable;
        }


        private string FormatPhoneNumber(string phoneNumber)
        {
            if (string.IsNullOrEmpty(phoneNumber))
                return phoneNumber;

            var digits = Regex.Replace(phoneNumber, @"[^\d]", "");

            if (digits.Length == 11)
            {
                return $"(+55) {digits.Substring(0, 2)} {digits.Substring(2, 5)}-{digits.Substring(7, 4)}";
            }
            else if (digits.Length == 10)
            {
                return $"(+55) {digits.Substring(0, 2)} {digits.Substring(2, 4)}-{digits.Substring(6, 4)}";
            }

            return phoneNumber;
        }

        private void SaveToExcel(DataTable dataTable, string filepath)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Dados");
                worksheet.Cells["A1"].LoadFromDataTable(dataTable, true);

                var dataCadastroColumn = worksheet.Cells["B2:B" + (dataTable.Rows.Count + 1)];
                dataCadastroColumn.Style.Numberformat.Format = "dd/MM/yyyy";

                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                package.SaveAs(new FileInfo(filepath));
            }
        }
    }
}