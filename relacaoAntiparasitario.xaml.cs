using ClosedXML.Excel;
using System.Data;
using System.Windows;
using System.Windows.Controls;

namespace CRM
{
    public partial class relacaoAntiparasitario : Window
    {
        public relacaoAntiparasitario()
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
                    var outputFilePath = System.IO.Path.Combine(desktopPath, "Antiparasitario.xlsx");

                    // Carregar o arquivo Excel
                    var workbook = new XLWorkbook(inputFilePath);
                    var worksheet = workbook.Worksheet(1);
                    var dataTable = worksheet.RangeUsed().AsTable().AsNativeDataTable();

                    // Adicionar a coluna "Grupo" ao DataTable
                    dataTable.Columns.Add("Grupo", typeof(int)); 

                    // Filtrar registros com caracteres indesejados
                    var caracteresIndesejados = new List<string> { "@", "*", "#"};
                    var pattern = string.Join("|", caracteresIndesejados.Select(System.Text.RegularExpressions.Regex.Escape));

                    var filteredRows = dataTable.AsEnumerable()
                        .Where(row => !System.Text.RegularExpressions.Regex.IsMatch(row.Field<string>("nome"), pattern))
                        .ToList();


                    // Data atual
                    var dataAtual = DateTime.Today;

                    // Definir grupos de dias e produtos
                    var gruposProdutos = new Dictionary<int, List<int>>
                        {
                            { 14, new List<int> { 52333 } },
                            { 28, new List<int> { 3996 } },
                            { 30, new List<int> { 964, 725, 722, 11, 15, 493, 499, 494, 47211, 48769, 49689, 43919, 41404, 52799, 51370, 731, 43616, 47600  } },
                            { 35, new List<int> { 38103, 45158, 45156 } },
                            { 37, new List<int> { 52966, 52967, 52968, 52965, 52969 } },
                            { 48, new List<int> { 45640 } },
                            { 49, new List<int> { 50725 } },
                            { 84, new List<int> { 152, 39382, 39150,39740, 39772, 43785, 49816, 49739, 49738 } },
                            { 90, new List<int> { 41464, 4166, 4421, 4165, 4164, 4163 } },
                            { 105, new List<int> { 49720, 49722, 51354, 51949, 49721, 49723 } },
                            { 120, new List<int> { 318 } },
                            { 144, new List<int> { 45640 } },
                            { 150, new List<int> { 48778 } },
                            { 168, new List<int> { 310 } },
                            { 180, new List<int> { 2060, 498 } },
                            { 240, new List<int> { 320, 41533 } }
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
                        newRow["fone2"] = row["Fone2"];
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
