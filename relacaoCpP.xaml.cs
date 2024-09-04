using ClosedXML.Excel;
using CRM.conexao;
using System.Data;
using System.IO;
using System.Windows.Controls;
using System.Windows;

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
            dataInicial.SelectedDateChanged += OnDateSelected;
            dataFinal.SelectedDateChanged += OnDateSelected;
            txtProduto.TextChanged += OnTextChanged;
            btnSearch.Click += OnSearchClick;
            btnExportarExcel.Click += OnExportarExcelClick; // Adicione este evento
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
            DateTime startDate = dataInicial.SelectedDate.Value;
            DateTime endDate = dataFinal.SelectedDate.Value;
            string produtoFilter = txtProduto.Text.Trim();
            var termosBusca = produtoFilter.Split(new[] { '%' }, StringSplitOptions.RemoveEmptyEntries);
            var caracteresIndesejados = new List<string> { "@", "*", "#", "MERCADO LIVRE", "CONSUMIDOR FINAL" };
            dbHelper = new conexaoCpP();
            DataTable dt = dbHelper.FetchData(startDate, endDate);

            var filteredData = dt.AsEnumerable()
                                 .Where(row => termosBusca.All(termo => row.Field<string>("Nome_Produto")
                                 .IndexOf(termo, StringComparison.OrdinalIgnoreCase) >= 0) &&
                                 !caracteresIndesejados.Any(ci => row.Field<string>("nome")
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

        private void OnExportarExcelClick(object sender, RoutedEventArgs e)
        {
            if (listaProd.ItemsSource == null)
            {
                MessageBox.Show("Não há dados para exportar.", "Aviso", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Criar a pasta 'Relações' no desktop se não existir
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string relacoesPath = Path.Combine(desktopPath, "Relações");
            if (!Directory.Exists(relacoesPath))
            {
                Directory.CreateDirectory(relacoesPath);
            }

            // Caminho completo do arquivo
            string filePath = Path.Combine(relacoesPath, "RelatorioClientesPromo.xlsx");

            // Criar o arquivo Excel
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Clientes");
                var dataTable = (DataView)listaProd.ItemsSource;

                // Adicionar os dados da DataTable à planilha
                worksheet.Cell(1, 1).InsertTable(dataTable.ToTable(), "Clientes");

                // Salvar o arquivo
                workbook.SaveAs(filePath);
            }

            MessageBox.Show("Arquivo exportado com sucesso!", "Exportar Excel", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}
