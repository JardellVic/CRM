using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace CRM
{
    public partial class MappingWindow : Window
    {
        private readonly List<string> columnNames;
        private readonly List<List<string>> rowData;
        private readonly Dictionary<string, int> columnIndexMap;

        public string ColunaNumeroSelecionada { get; set; }
        public string ColunaNomeSelecionada { get; set; }
        public string ColunaVariaveisSelecionada { get; set; }
        public List<string> DadosSelecionados { get; private set; }

        public MappingWindow(int paramsCount, List<string> columnNames, List<List<string>> rowData)
        {
            InitializeComponent();
            this.columnNames = columnNames;
            this.rowData = rowData;
            this.columnIndexMap = CreateColumnIndexMap(columnNames);
            CreateRadioButtons(paramsCount, columnNames.Count);
        }

        private Dictionary<string, int> CreateColumnIndexMap(List<string> columnNames)
        {
            var map = new Dictionary<string, int>();
            for (int i = 0; i < columnNames.Count; i++)
            {
                map[columnNames[i]] = i;
            }
            return map;
        }

        private void CreateRadioButtons(int paramsCount, int columnCount)
        {
            ColumnMappingPanel.Items.Clear();

            int foneColumnIndex = columnIndexMap["fone"];

            for (int i = 0; i < columnCount; i++)
            {
                if (i == foneColumnIndex)
                {
                    ColunaNumeroSelecionada = columnNames[i]; // Define a coluna "Número"
                    continue;
                }

                var grid = new Grid
                {
                    Margin = new Thickness(5)
                };

                grid.ColumnDefinitions.Add(new ColumnDefinition()); // Para o label da coluna
                grid.ColumnDefinitions.Add(new ColumnDefinition()); // Para o RadioButton "Número"

                for (int j = 0; j < paramsCount; j++)
                {
                    grid.ColumnDefinitions.Add(new ColumnDefinition());
                }

                var columnLabel = new TextBlock
                {
                    Text = columnNames[i],
                    Margin = new Thickness(5),
                    Foreground = System.Windows.Media.Brushes.White,
                    FontSize = 16
                };
                Grid.SetColumn(columnLabel, 0);
                grid.Children.Add(columnLabel);

                var numberRadioButton = new RadioButton
                {
                    Content = "Número",
                    Margin = new Thickness(5),
                    Foreground = System.Windows.Media.Brushes.White,
                    FontSize = 16,
                    Visibility = Visibility.Collapsed
                };
                Grid.SetColumn(numberRadioButton, 1);
                grid.Children.Add(numberRadioButton);

                for (int j = 1; j <= paramsCount; j++)
                {
                    var paramRadioButton = new RadioButton
                    {
                        Content = "Var" + j,
                        Margin = new Thickness(5),
                        Foreground = System.Windows.Media.Brushes.White,
                        FontSize = 16
                    };
                    Grid.SetColumn(paramRadioButton, j + 1);
                    grid.Children.Add(paramRadioButton);
                }

                ColumnMappingPanel.Items.Add(grid);
            }

            var nomeColunaIndex = columnIndexMap.ContainsKey("nome") ? columnIndexMap["nome"] : -1;
            ColunaNomeSelecionada = nomeColunaIndex != -1 ? columnNames[nomeColunaIndex] : string.Empty;
        }

        private void ConfirmButton_Click(object sender, RoutedEventArgs e)
        {
            DadosSelecionados = new List<string>();
            var variaveis = new List<string>();

            // Mapeia cada variável para sua coluna
            var selectedVariables = new Dictionary<int, string>();

            foreach (Grid grid in ColumnMappingPanel.Items)
            {
                var columnLabel = (TextBlock)grid.Children[0];
                var columnName = columnLabel.Text;
                int columnIndex = columnIndexMap[columnName];

                foreach (var child in grid.Children)
                {
                    if (child is RadioButton radioButton && radioButton.IsChecked == true)
                    {
                        var selectedValue = rowData[0][columnIndex];

                        if (radioButton.Content.ToString() == "Número")
                        {
                            ColunaNumeroSelecionada = columnName; // Atualiza a coluna número selecionada
                        }
                        else if (radioButton.Content.ToString().StartsWith("Var"))
                        {
                            int varIndex = int.Parse(radioButton.Content.ToString().Substring(3)) - 1;
                            selectedVariables[varIndex] = columnName; // Mapeia o índice da variável para a coluna
                        }
                        break;
                    }
                }
            }

            // Ordena as variáveis pela sua posição e converte para JSON
            var sortedVariables = new List<string>();
            for (int i = 0; i < selectedVariables.Count; i++)
            {
                if (selectedVariables.TryGetValue(i, out string columnName))
                {
                    sortedVariables.Add(columnName);
                }
            }

            ColunaVariaveisSelecionada = "[" + string.Join(",", sortedVariables.ConvertAll(v => $"\"{v}\"")) + "]";


            this.Close();
        }

        private List<List<string>> GetRowDataFromExcel(string filePath)
        {
            var rowData = new List<List<string>>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];

                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    var rowValues = new List<string>();

                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        rowValues.Add(worksheet.Cells[row, col].Text);
                    }

                    rowData.Add(rowValues);
                }
            }

            return rowData;
        }

        private void btnCancelar_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
