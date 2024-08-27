using OfficeOpenXml;
using CRM.conexao;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace CRM
{
    public partial class relatorioRetorno : Window
    {
        private conexaoMouraRetorno dbHelper;
        private DataTable _dataTable;

        public relatorioRetorno()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            dbHelper = new conexaoMouraRetorno();
        }

        private void txtClientes_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                string codigo = txtClientes.Text.Trim();

                if (!string.IsNullOrEmpty(codigo))
                {
                    lstClientes.Items.Add(codigo);
                    txtClientes.Clear();
                }

                e.Handled = true;
            }
        }

        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            var codigos = lstClientes.Items.OfType<string>().ToList();

            if (codigos.Count == 0)
            {
                MessageBox.Show("Por favor, adicione pelo menos um código de cliente.", "Aviso", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!dataInicial.SelectedDate.HasValue || !dateFinal.SelectedDate.HasValue)
            {
                MessageBox.Show("Por favor, selecione as datas de início e fim.", "Aviso", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            DateTime startDate = dataInicial.SelectedDate.Value;
            DateTime endDate = dateFinal.SelectedDate.Value;

            try
            {
                _dataTable = dbHelper.FetchData(codigos, startDate, endDate);

                if (_dataTable.Rows.Count > 0)
                {
                    MessageBox.Show($"Foram encontrados {_dataTable.Rows.Count} registros.", "Resultados", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show("Nenhum registro encontrado.", "Resultados", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao buscar os dados: {ex.Message}", "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void btnExportarExcel_Click(object sender, RoutedEventArgs e)
        {
            if (_dataTable == null || _dataTable.Rows.Count == 0)
            {
                MessageBox.Show("Nenhum dado para exportar.", "Aviso", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Relatório");

                // Adicionar cabeçalhos
                worksheet.Cells[1, 1].Value = "Código";
                worksheet.Cells[1, 2].Value = "Nome";
                worksheet.Cells[1, 3].Value = "Telefone";
                worksheet.Cells[1, 4].Value = "Telefone 2";
                worksheet.Cells[1, 5].Value = "Produto";
                worksheet.Cells[1, 6].Value = "Nome do Produto";
                worksheet.Cells[1, 7].Value = "Data da Venda";
                worksheet.Cells[1, 8].Value = "Quantidade do Item";
                worksheet.Cells[1, 9].Value = "Valor Total do Item";
                worksheet.Cells[1, 10].Value = "Empresa";
                worksheet.Cells[1, 11].Value = "Vendedor";

                int row = 2;
                foreach (DataRow dataRow in _dataTable.Rows)
                {
                    worksheet.Cells[row, 1].Value = dataRow["codigo"];
                    worksheet.Cells[row, 2].Value = dataRow["nome"];
                    worksheet.Cells[row, 3].Value = dataRow["fone"];
                    worksheet.Cells[row, 4].Value = dataRow["fone2"];
                    worksheet.Cells[row, 5].Value = dataRow["Produto"];
                    worksheet.Cells[row, 6].Value = dataRow["Nome_Produto"];
                    worksheet.Cells[row, 7].Value = dataRow["Data da Venda"];
                    worksheet.Cells[row, 8].Value = dataRow["Quantidade do Item"];
                    worksheet.Cells[row, 9].Value = dataRow["Valor Total do Item"];
                    worksheet.Cells[row, 10].Value = dataRow["Empresa"];
                    worksheet.Cells[row, 11].Value = dataRow["Vendedor"];
                    row++;
                }

                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                string filePath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Relatorio.xlsx");
                package.SaveAs(new System.IO.FileInfo(filePath));

                MessageBox.Show($"Os dados foram exportados com sucesso para {filePath}.", "Exportação Concluída", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
    }
}
