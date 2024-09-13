using ClosedXML.Excel;
using CRM.conexao;
using System.Data;
using System.IO;
using System.Windows.Controls;
using System.Windows;
using System.Text.RegularExpressions;

namespace CRM
{
    public partial class relacaoCpP : Window
    {
        private conexaoCpP dbHelper;

        public relacaoCpP()
        {
            InitializeComponent();
            txtProduto.IsEnabled = false;
            btnSearch.IsEnabled = false;
            dataInicial.SelectedDateChanged += OnDateSelected!;
            dataFinal.SelectedDateChanged += OnDateSelected!;
            txtProduto.TextChanged += OnTextChanged;
            btnSearch.Click += OnSearchClick;
            btnExportarExcel.Click += OnExportarExcelClick;
        }

        private void OnDateSelected(object sender, SelectionChangedEventArgs e)
        {
            if (dataInicial.SelectedDate.HasValue && dataFinal.SelectedDate.HasValue)
            {
                txtProduto.IsEnabled = true;
            }
            else
            {
                txtProduto.IsEnabled = false;
                btnSearch.IsEnabled = false;
            }
        }

        private void OnTextChanged(object sender, TextChangedEventArgs e)
        {
            btnSearch.IsEnabled = !string.IsNullOrWhiteSpace(txtProduto.Text);
        }

        private void OnSearchClick(object sender, RoutedEventArgs e)
        {
            DateTime startDate = dataInicial.SelectedDate!.Value;
            DateTime endDate = dataFinal.SelectedDate!.Value;
            string produtoFilter = txtProduto.Text.Trim();
            var termosBusca = produtoFilter.Split(new[] { '%' }, StringSplitOptions.RemoveEmptyEntries);
            var caracteresIndesejados = new List<string> { "@", "*", "#", "MERCADO LIVRE", "CONSUMIDOR FINAL" };
            dbHelper = new conexaoCpP();
            DataTable dt = dbHelper.FetchData(startDate, endDate);

            var filteredData = dt.AsEnumerable()
                                 .Where(row => termosBusca.All(termo => row.Field<string>("Nome_Produto")!
                                 .IndexOf(termo, StringComparison.OrdinalIgnoreCase) >= 0) &&
                                 !caracteresIndesejados.Any(ci => row.Field<string>("nome")!
                                 .IndexOf(ci, StringComparison.OrdinalIgnoreCase) >= 0));

            if (filteredData.Any())
            {
                listaProd.ItemsSource = filteredData.CopyToDataTable().DefaultView;
            }
            else
            {
                listaProd.ItemsSource = null;
            }
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

        private void OnExportarExcelClick(object sender, RoutedEventArgs e)
        {
            if (listaProd.ItemsSource == null)
            {
                MessageBox.Show("Não há dados para exportar.", "Aviso", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string relacoesPath = Path.Combine(desktopPath, "Relações");
            if (!Directory.Exists(relacoesPath))
            {
                Directory.CreateDirectory(relacoesPath);
            }

            string filePath = Path.Combine(relacoesPath, "RelatorioClientesPromo.xlsx");

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Clientes");
                var dataView = (DataView)listaProd.ItemsSource;
                var dataTable = dataView.ToTable();

                foreach (DataRow row in dataTable.Rows)
                {
                    if (dataTable.Columns.Contains("fone"))
                    {
                        row["fone"] = FormatPhoneNumber(row["fone"].ToString()!);
                    }

                    if (dataTable.Columns.Contains("fone2"))
                    {
                        row["fone2"] = FormatPhoneNumber(row["fone2"].ToString()!);
                    }
                }

                worksheet.Cell(1, 1).InsertTable(dataTable, "Clientes");

                workbook.SaveAs(filePath);
            }

            MessageBox.Show("Arquivo exportado com sucesso!", "Exportar Excel", MessageBoxButton.OK, MessageBoxImage.Information);
        }

    }
}
