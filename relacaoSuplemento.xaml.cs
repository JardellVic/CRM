using ClosedXML.Excel;
using System.Data;
using System.Windows;
using System.Windows.Controls;

namespace CRM
{
    public partial class relacaoSuplemento : Window
    {
        public relacaoSuplemento()
        {
            InitializeComponent();
            StartProcessing();
        }

        private async void StartProcessing()
        {
            ProgressBar.IsIndeterminate = true;
            await ProcessarDados();
            ProgressBar.IsIndeterminate = false;
        }

        private async Task ProcessarDados()
        {
            await Task.Run(() =>
            {
                try
                {
                    // Obtém o caminho da área de trabalho do usuário
                    var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                    // Caminho dos arquivos na área de trabalho
                    var inputFilePath = System.IO.Path.Combine(desktopPath, "Banco.xlsx");
                    var outputFilePath = System.IO.Path.Combine(desktopPath, "Suplemento.xlsx");

                    // Carregar o arquivo Excel
                    var workbook = new XLWorkbook(inputFilePath);
                    var worksheet = workbook.Worksheet(1);
                    var dataTable = worksheet.RangeUsed().AsTable().AsNativeDataTable();

                    // Adicionar a coluna "Grupo" ao DataTable
                    dataTable.Columns.Add("Grupo", typeof(int));

                    // Filtrar registros com caracteres indesejados
                    var caracteresIndesejados = new List<string> { "@", "*", "#", "MERCADO LIVRE", "CONSUMIDOR FINAL" };
                    var pattern = string.Join("|", caracteresIndesejados.Select(System.Text.RegularExpressions.Regex.Escape));

                    var filteredRows = dataTable.AsEnumerable()
                        .Where(row => !System.Text.RegularExpressions.Regex.IsMatch(row.Field<string>("nome"), pattern))
                        .ToList();


                    // Data atual
                    var dataAtual = DateTime.Today;

                    // Definir grupos de dias e produtos
                    var gruposProdutos = new Dictionary<int, List<int>>
                        {
                            { 19, new List<int> { 46050, 44229 } },
                            { 29, new List<int> { 45288, 769, 745, 44188, 754, 41644, 35, 524, 522, 40969, 45288, 45049, 44503, 43299, 43298, 44899, 40707, 43284, 2982, 43550, 45049, 42780, 41991, 42868, 45080, 51517, 49510, 49511, 47545, 40931, 53434, 47578, 49537, 42421, 53379, 629, 916, 38283, 185, 38742, 511, 935, 49908, 50778, 51517, 49500, 49501, 49503, 47545, 49476, 47545, 49476,47546, 53434, 53435, 47578, 49533, 49536, 49537, 42421, 632, 630, 631, 629, 627, 628, 38292, 915, 38281, 916, 38283, 917, 38288, 185, 634, 38742, 511, 935, 936, 51531, 42652, 46042, 46043, 49678, 47196, 50816, 49513, 44273, 774, 775, 780, 370, 45754, 49677, 51350, 43548, 43549, 43547, 51664, 43552, 51346, 43546, 52504, 52633, 43550, 43551, 957, 46041, 956, 77  } },
                            { 39, new List<int> { 51531, 42652, 46042, 49678, 47196 } },
                            { 59, new List<int> { 44273, 774, 780, 49677, 51350, 43548, 49580, 44921 } },
                        };

                    // Lista para armazenar todos os resultados
                    var resultadosCompletos = new List<DataRow>();

                    // Filtrar os produtos por grupos de dias
                    foreach (var grupo in gruposProdutos)
                    {
                        var dias = grupo.Key;
                        var produtos = grupo.Value;
                        var dataFiltro = dataAtual.AddDays(-dias).ToString("dd/MM/yyyy");

                        // Filtragem das linhas
                        var dfFiltrado = filteredRows
                            .Where(row =>
                            {
                                var produto = row.Field<object>("Produto");
                                int produtoInt;

                                // Tentar converter para int, se falhar, usar um valor padrão ou pular a linha
                                if (produto != null && int.TryParse(produto.ToString(), out produtoInt))
                                {
                                    return row.Field<string>("Data da Venda") == dataFiltro &&
                                           produtos.Contains(produtoInt);
                                }

                                return false;
                            })
                            .ToList();

                        foreach (var row in dfFiltrado)
                        {
                            var newRow = row;
                            newRow["Grupo"] = dias;
                            resultadosCompletos.Add(newRow);
                        }

                    }

                    // Criar uma nova DataTable apenas com as colunas desejadas
                    var resultadosFiltradosDataTable = new DataTable();
                    resultadosFiltradosDataTable.Columns.Add("nome", typeof(string));
                    resultadosFiltradosDataTable.Columns.Add("fone", typeof(string));
                    resultadosFiltradosDataTable.Columns.Add("fone2", typeof(string));
                    resultadosFiltradosDataTable.Columns.Add("Nome_Produto", typeof(string));
                    resultadosFiltradosDataTable.Columns.Add("Data da Venda", typeof(string));
                    resultadosFiltradosDataTable.Columns.Add("Grupo", typeof(int));

                    // Copiar as linhas filtradas para a nova DataTable
                    foreach (var row in resultadosCompletos)
                    {
                        var newRow = resultadosFiltradosDataTable.NewRow();
                        newRow["nome"] = row["Nome"];
                        newRow["fone"] = row["Fone"];
                        newRow["fone2"] = row["fone2"];
                        newRow["Nome_Produto"] = row["Nome_Produto"];
                        newRow["Data da Venda"] = row["Data da Venda"];
                        newRow["Grupo"] = row["Grupo"];
                        resultadosFiltradosDataTable.Rows.Add(newRow);
                    }

                    var newWorkbook = new XLWorkbook();
                    var newWorksheet = newWorkbook.Worksheets.Add("Resultado");
                    newWorksheet.Cell(1, 1).InsertTable(resultadosFiltradosDataTable);
                    newWorkbook.SaveAs(outputFilePath);

                    Dispatcher.Invoke(() =>
                    {
                        MessageBox.Show($"Arquivo salvo: {outputFilePath}", "Concluído", MessageBoxButton.OK);
                        this.Close();
                    });
                }
                catch (Exception ex)
                {
                    Dispatcher.Invoke(() =>
                    {
                        MessageBox.Show($"Erro: {ex.Message}", "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
                        this.Close();
                    });
                }
            });
        }



        private DataTable ConvertListToDataTable(List<DataRow> rows)
        {
            if (rows.Count == 0)
                return new DataTable();

            var dataTable = rows[0].Table.Clone();
            foreach (var row in rows)
            {
                dataTable.ImportRow(row);
            }
            return dataTable;
        }
    }
}
