using ClosedXML.Excel;
using System.Data;
using System.Windows;
using System.Windows.Controls;

namespace CRM
{
    public partial class relacaoMilteforan : Window
    {
        public relacaoMilteforan()
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
                    var outputFilePath = System.IO.Path.Combine(desktopPath, "Milteforan.xlsx");

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
                    var isSegundaFeira = dataAtual.DayOfWeek == DayOfWeek.Monday;

                    // Definir grupos de dias e produtos
                    var gruposProdutos = new Dictionary<int, List<int>>
                        {
                            { 19, new List<int> { 52836 } },
                            { 29, new List<int> { 45288, 769, 745, 44188, 754, 45613, 41644 } },
                            { 39, new List<int> { 51531, 42652 } },
                            { 59, new List<int> { 50816, 44273, 774, 780 } }
                    };

                    // Lista para armazenar todos os resultados
                    var resultadosCompletos = new List<DataRow>();

                    // Filtrar os produtos por grupos de dias
                    foreach (var grupo in gruposProdutos)
                    {
                        var dias = grupo.Key;
                        var produtos = grupo.Value;

                        var datasFiltro = new List<string>
                            {
                            dataAtual.AddDays(-dias).ToString("dd/MM/yyyy")
                        };

                        if (isSegundaFeira)
                        {
                            datasFiltro.Add(dataAtual.AddDays(-dias - 1).ToString("dd/MM/yyyy"));
                            datasFiltro.Add(dataAtual.AddDays(-dias - 2).ToString("dd/MM/yyyy"));
                            datasFiltro.Add(dataAtual.AddDays(-dias - 3).ToString("dd/MM/yyyy"));
                        }

                        foreach (var dataFiltro in datasFiltro)
                        {
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
