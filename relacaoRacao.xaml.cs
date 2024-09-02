using OfficeOpenXml;
using System.Data;
using System.IO;
using System.Windows;
using System.Windows.Threading;

namespace CRM
{
    public partial class relacaoRacao : Window
    {
        public relacaoRacao()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            StartProcessing();
        }

        private async void StartProcessing()
        {
            ProgressBar.IsIndeterminate = true;
            await Task.Run(() => ProcessData());
            ProgressBar.IsIndeterminate = false;

            MessageBox.Show("Arquivo salvo com sucesso!", "Sucesso", MessageBoxButton.OK, MessageBoxImage.Information);
            this.Close();
        }

        private void ProcessData()
{
    try
    {
        var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        var inputFilePath = Path.Combine(desktopPath, "Banco.xlsx");
        var outputFilePath = Path.Combine(desktopPath, "Racao.xlsx");

        // Verifique se o arquivo de saída já existe e exclua-o, se necessário
        if (File.Exists(outputFilePath))
        {
            File.Delete(outputFilePath);
        }

        using (var package = new ExcelPackage(new FileInfo(inputFilePath)))
        {
            var worksheet = package.Workbook.Worksheets.FirstOrDefault();
            if (worksheet == null) return;

            var table = new DataTable();

            foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
            {
                table.Columns.Add(firstRowCell.Text);
            }

            for (var rowNumber = 2; rowNumber <= worksheet.Dimension.End.Row; rowNumber++)
            {
                var row = worksheet.Cells[rowNumber, 1, rowNumber, worksheet.Dimension.End.Column];
                var newRow = table.NewRow();
                foreach (var cell in row)
                {
                    newRow[cell.Start.Column - 1] = cell.Text;
                }
                table.Rows.Add(newRow);
            }

            // Filtrar produtos que começam com "RAÇÃO" ou "RACAO" e clientes sem caracteres especiais
            var recentDate = DateTime.Now.AddMonths(-6);

            var filteredRows = table.AsEnumerable()
                .Where(row => !row["nome"].ToString().Contains("#") &&
                              !row["nome"].ToString().Contains("@") &&
                              !row["nome"].ToString().Contains("&") &&
                              !row["nome"].ToString().Contains("MERCADO LIVRE") &&
                              !row["nome"].ToString().Contains("CONSUMIDOR FINAL") &&
                              DateTime.TryParse(row["Data da Venda"].ToString(), out DateTime dataVenda) && dataVenda >= recentDate &&
                              (row["Nome_Produto"].ToString().StartsWith("RAÇÃO") || row["Nome_Produto"].ToString().StartsWith("RACAO")))
                .CopyToDataTable();

            // Calcular a média de dias entre compras por cliente
            var groupedRows = filteredRows.AsEnumerable()
                .GroupBy(row => row["nome"].ToString())
                .Select(g =>
                {
                    var orderedDates = g.Select(row => DateTime.Parse(row["Data da Venda"].ToString())).OrderBy(d => d).ToList();
                    var mediaDias = orderedDates.Count > 1 ? Math.Round(orderedDates.Zip(orderedDates.Skip(1), (a, b) => (b - a).TotalDays).Average()) : 0;
                    var dataMax = orderedDates.Max();
                    var proximaCompra = dataMax.AddDays(mediaDias);

                    return new
                    {
                        Codigo = g.First()["codigo"].ToString(),
                        Nome = g.Key,
                        Fone = g.First()["fone"].ToString(),
                        Fone2 = g.First()["fone2"].ToString(),
                        NomeProduto = g.First()["Nome_Produto"].ToString(), // Obtém o valor de "Nome_Produto"
                        MediaDiasEntreCompras = mediaDias,
                        DataUltimaCompra = dataMax,
                        ProximaCompra = proximaCompra
                    };
                })
                .Where(x => x.ProximaCompra.Date == DateTime.Now.AddDays(3).Date)
                .ToList();

            // Criar DataTable para salvar no Excel
            var resultTable = new DataTable();
            resultTable.Columns.Add("Nome", typeof(string));
            resultTable.Columns.Add("Fone", typeof(string));
            resultTable.Columns.Add("Fone2", typeof(string));
            resultTable.Columns.Add("Nome_Produto", typeof(string)); // Alterado para "Nome_Produto"
            resultTable.Columns.Add("Media Dias Entre Compras", typeof(double));
            resultTable.Columns.Add("Data Última Compra", typeof(DateTime));
            resultTable.Columns.Add("Próxima Compra", typeof(DateTime));

            foreach (var item in groupedRows)
            {
                var row = resultTable.NewRow();
                row["Nome"] = item.Nome;
                row["Fone"] = item.Fone;
                row["Fone2"] = item.Fone2;
                row["Nome_Produto"] = item.NomeProduto; // Usando o valor de "Nome_Produto"
                row["Media Dias Entre Compras"] = item.MediaDiasEntreCompras;
                row["Data Última Compra"] = item.DataUltimaCompra;
                row["Próxima Compra"] = item.ProximaCompra;
                resultTable.Rows.Add(row);
            }

            // Salvar os resultados em um novo arquivo Excel
            using (var newPackage = new ExcelPackage(new FileInfo(outputFilePath)))
            {
                var newWorksheet = newPackage.Workbook.Worksheets.Add("Resultado");
                newWorksheet.Cells["A1"].LoadFromDataTable(resultTable, true);

                // Aplicar formatação para as colunas de data
                var dataUltimaCompraCol = newWorksheet.Cells[2, resultTable.Columns["Data Última Compra"].Ordinal + 1, resultTable.Rows.Count + 1, resultTable.Columns["Data Última Compra"].Ordinal + 1];
                var proximaCompraCol = newWorksheet.Cells[2, resultTable.Columns["Próxima Compra"].Ordinal + 1, resultTable.Rows.Count + 1, resultTable.Columns["Próxima Compra"].Ordinal + 1];

                dataUltimaCompraCol.Style.Numberformat.Format = "dd/MM/yyyy";
                proximaCompraCol.Style.Numberformat.Format = "dd/MM/yyyy";

                newPackage.Save();
            }
        }
    }
    catch (Exception ex)
    {
        Dispatcher.Invoke(() =>
        {
            MessageBox.Show($"Erro ao processar o arquivo: {ex.Message}", "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
        });
    }
}



    }
}
