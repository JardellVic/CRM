using ClosedXML.Excel;
using OfficeOpenXml;
using CRM.conexao;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Windows;

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
                // Converter a data de acordo com o formato dd/MM/yyyy
                DateTime startDate = DateTime.ParseExact(txtDataInicial.Text, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture).Date;
                //DateTime endDate = DateTime.Now.AddDays(-1).Date.AddSeconds(-1);
                DateTime endDate = startDate.AddDays(1).AddSeconds(-1); // Fim do dia

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
            string outputPath = Path.Combine(desktopPath, "Welcome.xlsx");
            string? directoryPath = Path.GetDirectoryName(outputPath);

            if (directoryPath != null)
            {
                Directory.CreateDirectory(directoryPath);
            }
            else
            {
                // Caso o caminho do diretório seja nulo, exibe uma mensagem de erro
                MessageBox.Show("Não foi possível determinar o caminho do diretório para salvar o arquivo.", "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // Filtra as colunas desejadas
            DataTable filteredTable = FilterColumns(_dataTable);

            // Exibe a barra de progresso e define como indeterminada
            progressBar.Visibility = Visibility.Visible;
            progressBar.IsIndeterminate = true;

            // Executa a operação de salvamento em uma tarefa separada
            await Task.Run(() => SaveToExcel(filteredTable, outputPath));

            // Oculta a barra de progresso
            progressBar.Visibility = Visibility.Collapsed;

            // Exibe uma mensagem de sucesso e fecha a janela
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

            // Adicionar as colunas específicas
            filteredTable.Columns.Add("Nome", typeof(string));
            filteredTable.Columns.Add("Data_Cadastro", typeof(DateTime));
            filteredTable.Columns.Add("Fone", typeof(string));
            filteredTable.Columns.Add("Fone2", typeof(string));

            // Preencher o DataTable filtrado com os dados
            foreach (DataRow row in dataTable.Rows)
            {
                DataRow newRow = filteredTable.NewRow();
                newRow["Nome"] = row["Nome"];
                newRow["Data_Cadastro"] = row["Data_Cadastro"];
                newRow["Fone"] = row["Fone"];
                newRow["Fone2"] = row["Fone2"];
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

                // Formata a coluna de Data_Cadastro no formato dd/MM/yyyy
                var dataCadastroColumn = worksheet.Cells["B2:B" + (dataTable.Rows.Count + 1)];
                dataCadastroColumn.Style.Numberformat.Format = "dd/MM/yyyy";

                // Ajusta o comprimento das colunas
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                package.SaveAs(new FileInfo(filepath));
            }
        }
    }
}