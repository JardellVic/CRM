using OfficeOpenXml;
using CRM.conexao;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Windows;
using System.Text.RegularExpressions;

namespace CRM
{
    public partial class relacaoVacina : Window
    {
        private DataTable _dataTable;
        private conexaoMouraVacina dbHelper;
        public relacaoVacina()
        {
            InitializeComponent();
            SetupDates();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // Inicializa o helper do banco de dados
            dbHelper = new conexaoMouraVacina();
        }

        private void SetupDates()
        {
            DateTime today = DateTime.Now;
            DateTime startDate = today.AddMonths(-11).AddDays(-28);
            txtDataInicial.Text = startDate.ToString("dd/MM/yyyy");
        }

        private async void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Converter a string para DateTime usando o formato específico "dd/MM/yyyy"
                DateTime startDate = DateTime.ParseExact(txtDataInicial.Text, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                DateTime endDate = startDate;

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

        private string FormatPhoneNumber(string phoneNumber)
        {
            if (string.IsNullOrEmpty(phoneNumber))
                return phoneNumber;

            // Remove non-numeric characters
            var digits = Regex.Replace(phoneNumber, @"[^\d]", "");

            // Format the string
            if (digits.Length == 11) // Format as +55 xx xxxxx-xxxx
            {
                return $"(+55) {digits.Substring(0, 2)} {digits.Substring(2, 5)}-{digits.Substring(7, 4)}";
            }
            else if (digits.Length == 10) // Format as +55 xx xxxx-xxxx
            {
                return $"(+55) {digits.Substring(0, 2)} {digits.Substring(2, 4)}-{digits.Substring(6, 4)}";
            }

            // If the number doesn't fit the pattern, return as is
            return phoneNumber;
        }

        private async void btnExportarExcel_Click(object sender, RoutedEventArgs e)
        {
            if (_dataTable == null || _dataTable.Rows.Count == 0)
            {
                MessageBox.Show("Nenhum dado para exportar.");
                return;
            }
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string outputPath = Path.Combine(desktopPath, "Vacina.xlsx");
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

            // Filtrar as colunas desejadas
            DataTable filteredTable = FilterColumns(_dataTable);

            progressBar.Visibility = Visibility.Visible;
            progressBar.IsIndeterminate = true;

            await Task.Run(() => SaveToExcel(filteredTable, outputPath));

            progressBar.Visibility = Visibility.Collapsed;

            MessageBox.Show($"Arquivo salvo!", "Concluído", MessageBoxButton.OK);
            this.Close();
        }

        private DataTable FilterColumns(DataTable dataTable)
        {
            DataTable filteredTable = new DataTable();

            // Adicionar as colunas específicas com os nomes corretos
            filteredTable.Columns.Add("nome", typeof(string));
            filteredTable.Columns.Add("Data", typeof(string));
            filteredTable.Columns.Add("fone", typeof(string));
            filteredTable.Columns.Add("Pet", typeof(string));
            filteredTable.Columns.Add("Serviço", typeof(string));

            // Filtrar as linhas com base nas condições especificadas
            var filteredRows = dataTable.AsEnumerable()
                .Where(row => !row["Proprietario"].ToString().Contains("#") &&
                              !row["Proprietario"].ToString().Contains("@") &&
                              !row["Proprietario"].ToString().Contains("&") &&
                              !row["Proprietario"].ToString().Contains("MERCADO LIVRE") &&
                              !row["Proprietario"].ToString().Contains("CONSUMIDOR FINAL"));

            // Preencher o DataTable filtrado com os dados filtrados
            foreach (var row in filteredRows)
            {
                DataRow newRow = filteredTable.NewRow();
                newRow["nome"] = row["Proprietario"];
                // Formatando a data para o formato dd/MM/yyyy
                newRow["Data"] = Convert.ToDateTime(row["data"]).ToString("dd/MM/yyyy");
                newRow["fone"] = FormatPhoneNumber(row["fone"].ToString());
                newRow["Pet"] = row["nome_animal"];
                newRow["Serviço"] = row["Servico"];
                filteredTable.Rows.Add(newRow);
            }

            return filteredTable;
        }

        private void SaveToExcel(DataTable dataTable, string filepath)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Dados");
                worksheet.Cells["A1"].LoadFromDataTable(dataTable, true);

                // Ajusta o comprimento das colunas
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                package.SaveAs(new FileInfo(filepath));
            }
        }
    }
}