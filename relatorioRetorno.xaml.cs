using OfficeOpenXml;
using CRM.conexao;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.IO;

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
            var predefinedCodes = new[] { "467", "497", "543", "429", "552", "486", "542", "506", "544" };
            foreach (var code in predefinedCodes)
            {
                lstClientes.Items.Add(code);
            }
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

        private void LstCod_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (lstClientes.SelectedItem != null)
            {
                lstClientes.Items.Remove(lstClientes.SelectedItem);
            }
        }

        private async void btnSearch_Click(object sender, RoutedEventArgs e)
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

            progressBar.IsIndeterminate = true;

            try
            {
                _dataTable = await Task.Run(() => dbHelper.FetchData(codigos, startDate, endDate));

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
            finally
            {
                progressBar.IsIndeterminate = false;
            }
        }

        private async void btnExportarExcel_Click(object sender, RoutedEventArgs e)
        {
            if (_dataTable == null || _dataTable.Rows.Count == 0)
            {
                MessageBox.Show("Nenhum dado para exportar.", "Aviso", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                progressBar.IsIndeterminate = true;

                await Task.Run(() =>
                {
                    using (var package = new ExcelPackage())
                    {
                        var worksheet = package.Workbook.Worksheets.Add("Relatório");

                        worksheet.Cells[1, 1].Value = "Cliente";
                        worksheet.Cells[1, 2].Value = "Produto";
                        worksheet.Cells[1, 3].Value = "Nome do Produto";
                        worksheet.Cells[1, 4].Value = "Data da Venda";
                        worksheet.Cells[1, 5].Value = "Vendedor";

                        int row = 2;
                        foreach (DataRow dataRow in _dataTable.Rows)
                        {
                            worksheet.Cells[row, 1].Value = dataRow["nome"];
                            worksheet.Cells[row, 2].Value = dataRow["Produto"];
                            worksheet.Cells[row, 3].Value = dataRow["Nome_Produto"];
                            worksheet.Cells[row, 4].Value = dataRow["Data da Venda"];
                            worksheet.Cells[row, 5].Value = dataRow["Vendedor"];
                            row++;
                        }


                        worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                        string relacoesPath = Path.Combine(desktopPath, "Relações");
                        if (!Directory.Exists(relacoesPath))
                        {
                            Directory.CreateDirectory(relacoesPath);
                        }

                        string filePath = Path.Combine(relacoesPath, "Relatorio Retorno.xlsx");
                        package.SaveAs(new System.IO.FileInfo(filePath));

                        Dispatcher.Invoke(() => MessageBox.Show($"Os dados foram exportados com sucesso para {filePath}.", "Exportação Concluída", MessageBoxButton.OK, MessageBoxImage.Information));
                    }
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao exportar os dados: {ex.Message}", "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                progressBar.IsIndeterminate = false;
            }
        }
    }
}