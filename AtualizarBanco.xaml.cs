using CRM.conexao;
using OfficeOpenXml;
using System;
using System.Data;
using System.IO;
using System.Threading.Tasks;
using System.Windows;

namespace CRM
{
    public partial class AtualizarBanco : Window
    {
        private DataTable dataTable;
        private conexaoMouraBanco dbHelper;

        public AtualizarBanco()
        {
            InitializeComponent();

            // Define o contexto da licença EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Inicializa o dataTable para evitar referências nulas
            dataTable = new DataTable();

            // Inicializa o helper do banco de dados
            dbHelper = new conexaoMouraBanco();

            // Define as datas predefinidas
            dataInicial.SelectedDate = DateTime.Now.AddMonths(-6).AddDays(-1); // 6 meses e 1 dia atrás
            dateFinal.SelectedDate = DateTime.Now.AddDays(-1); // Data de ontem
        }

        private async void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            if (dataInicial.SelectedDate.HasValue && dateFinal.SelectedDate.HasValue)
            {
                DateTime startDate = dataInicial.SelectedDate.Value;
                DateTime endDate = dateFinal.SelectedDate.Value;

                progressBar.IsIndeterminate = true;

                await Task.Run(() =>
                {
                    dataTable = dbHelper.FetchData(startDate, endDate);
                });

                progressBar.IsIndeterminate = false;
                MessageBox.Show($"Total de registros: {dataTable.Rows.Count}");
            }
            else
            {
                MessageBox.Show("Por favor, selecione as datas.");
            }
        }

        private void btnExportarExcel_Click(object sender, RoutedEventArgs e)
        {
            if (dataTable != null && dataTable.Rows.Count > 0)
            {
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string filePath = Path.Combine(desktopPath, "Banco.xlsx");

                using (ExcelPackage package = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Dados");

                    // Escrever DataTable no Excel
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        worksheet.Cells[1, i + 1].Value = dataTable.Columns[i].ColumnName;
                    }

                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataTable.Columns.Count; j++)
                        {
                            worksheet.Cells[i + 2, j + 1].Value = dataTable.Rows[i][j];
                        }
                    }

                    FileInfo fileInfo = new FileInfo(filePath);
                    package.SaveAs(fileInfo);
                }

                MessageBoxResult result = MessageBox.Show("Exportação concluída! Arquivo salvo na área de trabalho.", "Exportação Completa", MessageBoxButton.OK);

                if (result == MessageBoxResult.OK)
                {
                    this.Close(); // Fecha a janela
                }
            }
            else
            {
                MessageBox.Show("Nenhum dado para exportar. Execute a pesquisa primeiro.");
            }
        }
    }
}
