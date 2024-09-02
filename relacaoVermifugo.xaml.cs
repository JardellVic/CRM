﻿using ClosedXML.Excel;
using System.Data;
using System.Windows;
using System.Windows.Controls;

namespace CRM
{
    public partial class relacaoVermifugo : Window
    {
        public relacaoVermifugo()
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
                    var outputFilePath = System.IO.Path.Combine(desktopPath, "Vermifugo.xlsx");

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

                            { 29, new List<int> { 2033, 2966, 451 } },
                            { 89, new List<int> { 44862, 41015, 44420, 44837, 42341, 42342, 43327, 48485, 3130, 462, 38075, 2964, 677, 4578, 1919, 2004, 2881, 5163, 48485, 462, 690, 461, 448, 444, 48482, 45352, 5163, 5162, 690, 692} },
             
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
