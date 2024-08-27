using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using Microsoft.Win32;
using OfficeOpenXml;
using Newtonsoft.Json.Linq;
using CRM.conexao.API;

namespace CRM
{
    public class LineData
    {
        public string Numero { get; set; }
        public string Nome { get; set; }
        public List<string> Variaveis { get; set; }
    }

    public partial class Home : Window
    {
        public static Home Instance { get; private set; }
        public string TemplateIdSelecionado { get; set; }

        private readonly APIManager apiManager;
        private readonly HttpClient client;
        private Dictionary<string, string> templateTextMap;
        private Dictionary<string, int> templateParamsMap;
        private Dictionary<string, string> templateIdMap;

        public Home()
        {
            InitializeComponent();
            apiManager = new APIManager();
            client = new HttpClient();
            templateTextMap = new Dictionary<string, string>();
            templateParamsMap = new Dictionary<string, int>();
            templateIdMap = new Dictionary<string, string>();
            LoadTemplatesAsync();
            cmbTemplates.SelectionChanged += CmbTemplates_SelectionChanged;
            Instance = this;
            this.ResizeMode = ResizeMode.NoResize;
        }

        private async void LoadTemplatesAsync()
        {
            try
            {
                var templates = await apiManager.GetTemplatesAsync();
                templateTextMap.Clear();
                templateParamsMap.Clear();
                templateIdMap.Clear();
                cmbTemplates.Items.Clear();

                foreach (var template in templates)
                {
                    string nome = template.Key;
                    var (texto, paramsCount, id) = template.Value;
                    templateTextMap[nome] = texto;
                    templateParamsMap[nome] = paramsCount;
                    templateIdMap[nome] = id;
                    cmbTemplates.Items.Add(nome);
                }
            }
            catch (HttpRequestException e)
            {
                MessageBox.Show($"Erro ao acessar a API: {e.Message}", "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void AtualizarStatusBar(int quantidadeContatos)
        {
            statusContatos.Content = $"Quantidade de contatos: {quantidadeContatos}";

            double valorUtility = 0.008 * quantidadeContatos;
            statusUtility.Content = $"Valor Utility: ${valorUtility:F2}";

            double valorMarketing = 0.0625 * quantidadeContatos;
            statusMarketing.Content = $"Valor Marketing: ${valorMarketing:F2}";
        }

        private void CmbTemplates_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbTemplates.SelectedItem != null)
            {
                string selectedTemplate = cmbTemplates.SelectedItem.ToString();
                if (templateTextMap.ContainsKey(selectedTemplate))
                {
                    txtTemplate.Text = templateTextMap[selectedTemplate];
                    if (templateIdMap.ContainsKey(selectedTemplate))
                    {
                        TemplateIdSelecionado = templateIdMap[selectedTemplate];
                    }
                }
            }
        }

        private async void SelectFileButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                FilePathTextBox.Text = openFileDialog.FileName;

                if (cmbTemplates.SelectedItem != null)
                {
                    string selectedTemplate = cmbTemplates.SelectedItem.ToString();
                    int paramsCount = GetParamsCountForTemplate(selectedTemplate);

                    if (GetColumnCountFromExcel(openFileDialog.FileName) < paramsCount + 1)
                    {
                        ShowError($"O arquivo Excel deve ter pelo menos {paramsCount + 1} colunas.");
                        return;
                    }

                    OpenMappingWindow(paramsCount);
                }
            }
        }

        private void ShowError(string message)
        {
            MessageBox.Show(message, "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        private List<string> GetColumnNamesFromExcel(string filePath)
        {
            var columnNames = new List<string>();
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    columnNames.Add(worksheet.Cells[1, col].Text);
                }
            }
            return columnNames;
        }

        private List<List<string>> GetRowDataFromExcel(string filePath)
        {
            var rowData = new List<List<string>>();

            try
            {
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
            }
            catch (IOException ioEx)
            {
                MessageBox.Show($"Erro ao acessar o arquivo Excel: {ioEx.Message}", "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ocorreu um erro: {ex.Message}", "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            return rowData;
        }

        private void OpenMappingWindow(int paramsCount)
        {
            var columnNames = GetColumnNamesFromExcel(FilePathTextBox.Text);
            var rowData = GetRowDataFromExcel(FilePathTextBox.Text);

            // Atualiza a quantidade de contatos
            int quantidadeContatos = rowData.Count;
            AtualizarStatusBar(quantidadeContatos);

            MappingWindow mappingWindow = new MappingWindow(paramsCount, columnNames, rowData);
            mappingWindow.ShowDialog();

            // Processar dados da planilha conforme as colunas selecionadas
            if (!string.IsNullOrEmpty(mappingWindow.ColunaNumeroSelecionada) &&
                !string.IsNullOrEmpty(mappingWindow.ColunaNomeSelecionada))
            {
                ProcessRowData(mappingWindow.ColunaNumeroSelecionada, mappingWindow.ColunaNomeSelecionada, mappingWindow.ColunaVariaveisSelecionada, rowData);
            }
        }

        public List<LineData> LinhasParaEnviar { get; private set; }

        private void ProcessRowData(string colunaNumero, string colunaNome, string variaveisColuna, List<List<string>> rowData)
        {
            var linhas = new List<LineData>();
            var columnNames = GetColumnNamesFromExcel(FilePathTextBox.Text); // Certifique-se de que columnNames é acessível

            int numeroIndex = columnNames.IndexOf(colunaNumero);
            int nomeIndex = columnNames.IndexOf(colunaNome);
            List<int> variaveisIndices = variaveisColuna.Trim('[', ']').Split(',').Select(v => columnNames.IndexOf(v.Trim('"'))).ToList();

            foreach (var row in rowData)
            {
                if (row.Count > numeroIndex && row.Count > nomeIndex && variaveisIndices.All(i => i < row.Count))
                {
                    var lineData = new LineData
                    {
                        Numero = row[numeroIndex],
                        Nome = row[nomeIndex],
                        Variaveis = variaveisIndices.Select(i => row[i]).ToList()
                    };
                    linhas.Add(lineData);
                }
            }

            Home.Instance.LinhasParaEnviar = linhas;
        }

        private int GetParamsCountForTemplate(string templateName)
        {
            return templateParamsMap.ContainsKey(templateName) ? templateParamsMap[templateName] : 0;
        }

        private int GetColumnCountFromExcel(string filePath)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                return worksheet.Dimension.End.Column;
            }
        }

        private async void btnEnviarDisparo_Click(object sender, RoutedEventArgs e)
        {
            int sucessoCount = 0;
            int erroCount = 0;
            int totalLinhas = Home.Instance.LinhasParaEnviar.Count;

            // Inicializa a ProgressBar
            progressDisparo.IsIndeterminate = false;
            progressDisparo.Maximum = totalLinhas;
            progressDisparo.Value = 0;

            // Limpa o conteúdo anterior e inicializa o contador de progresso
            txtBlockConsole.Inlines.Clear();
            txtBlockConsole.Inlines.Add(new Run($"Iniciando envio... (0/{totalLinhas})") { Foreground = Brushes.Yellow });

            try
            {
                var linhasParaEnviar = Home.Instance.LinhasParaEnviar;

                foreach (var linha in linhasParaEnviar.Select((value, index) => new { value, index }))
                {
                    bool resultado = await EnviarLinhaAsync(linha.value);
                    if (resultado)
                    {
                        sucessoCount++;
                    }
                    else
                    {
                        erroCount++;
                        txtBlockConsole.Inlines.Add(new Run($"\nErro: {linha.value.Numero}") { Foreground = Brushes.Red });
                    }

                    // Atualiza o progresso e o contador
                    progressDisparo.Value = linha.index + 1;
                    txtBlockConsole.Inlines.Add(new Run($"\nProgresso: {linha.index + 1}/{totalLinhas}") { Foreground = Brushes.Blue });

                    await Task.Delay(1000); // Aguarda 1 segundo entre envios
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ocorreu um erro: {ex.Message}", "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                progressDisparo.IsIndeterminate = false; // Finaliza o progresso indeterminado
                MessageBox.Show($"Envios concluídos!\n\nSucessos: {sucessoCount}\nErros: {erroCount}",
                                "Resultado do Envio", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private async Task<bool> EnviarLinhaAsync(LineData linha)
        {
            try
            {
                // Chama o método para fazer o opt-in do número
                bool optinResult = await OptinNumeroAsync(linha.Numero);

                // Exibe a mensagem no console com base no resultado do opt-in
                if (optinResult)
                {
                    txtBlockConsole.Inlines.Add(new Run("\nOptIn feito com sucesso") { Foreground = Brushes.Green });
                }
                else
                {
                    txtBlockConsole.Inlines.Add(new Run("Erro ao fazer optin") { Foreground = Brushes.Red });
                    return false; // Se falhar no opt-in, não envia o template
                }

                var formData = new MultipartFormDataContent
                {
                    { new StringContent(TemplateIdSelecionado), "template_id" },
                    { new StringContent(linha.Numero), "numero" },
                    { new StringContent(linha.Nome), "nome_cliente" },
                    { new StringContent("[" + string.Join(",", linha.Variaveis.Select(v => $"\"{v}\"")) + "]"), "variaveis" },
                    { new StringContent("Pet"), "bot" },
                    { new StringContent("Inicio"), "menu_bot" }
                };

                HttpResponseMessage response = await client.PostAsync("http://18.230.12.44/api/v1/wpp/enviarTemplate?key=856adfb59d45471ab288e45d3e4d9a7865f9c075cc142", formData);
                string responseBody = await response.Content.ReadAsStringAsync();

                if (response.IsSuccessStatusCode)
                {
                    txtBlockConsole.Inlines.Add(new Run($"\nSucesso: {linha.Numero}") { Foreground = Brushes.Green });
                    return true;
                }
                else
                {
                    txtBlockConsole.Inlines.Add(new Run($"\nErro: {linha.Numero}") { Foreground = Brushes.Red });
                    return false;
                }
            }
            catch (Exception ex)
            {
                txtBlockConsole.Inlines.Add(new Run($"\nErro: {ex.Message}") { Foreground = Brushes.Red });
                return false;
            }
        }

        private async Task<bool> OptinNumeroAsync(string numero)
        {
            try
            {
                var formData = new MultipartFormDataContent
                {
                    { new StringContent("a86664b9-95de-4fd2-bc68-3b1e689d0a0f"), "app_id" },
                    { new StringContent(numero), "numero" },
                    { new StringContent("true"), "optin" }
                };

                HttpResponseMessage response = await client.PostAsync("http://whatsapp.petcaesecia.com.br/api/v1/wpp/alterarStatusOptinNumero?key=856adfb59d45471ab288e45d3e4d9a7865f9c075cc142", formData);
                string responseBody = await response.Content.ReadAsStringAsync();

                JObject jsonResponse = JObject.Parse(responseBody);
                string status = (string)jsonResponse["status"];
                string msg = (string)jsonResponse["data"]["msg"];

                // Checa se o status é "success" e a mensagem é a esperada
                return status == "success" && msg.Contains("A solicitação foi feita com sucesso");
            }
            catch (Exception ex)
            {
                // Em caso de erro, podemos retornar false ou lidar com a exceção conforme necessário
                return false;
            }
        }
    


//-----------------------------------------------------------------------------------------------------------------------------------------------------------
private void BancoMenuItem_Click(object sender, RoutedEventArgs e)
        {
            AtualizarBanco atualizarBancoPage = new AtualizarBanco();
            atualizarBancoPage.Show();
        }

        private void antiparasitario_Click(object sender, RoutedEventArgs e)
        {
            var relacaoWindow = new relacaoAntiparasitario();
            relacaoWindow.Show();
            lblGerarAntiparasitario.Text = "Gerar Anti-Parasitário:✅";
        }

        private void vrfcAntiparasitario_Click(object sender, RoutedEventArgs e)
        {
            // Obter o caminho do Desktop
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            // Combinar o caminho do Desktop com o nome do arquivo
            string filePath = Path.Combine(desktopPath, "Antiparasitario.xlsx");

            // Verificar se o arquivo existe
            if (File.Exists(filePath))
            {
                // Abrir o arquivo
                Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
            }
            else
            {
                MessageBox.Show("O arquivo Antiparasitario.xlsx não foi encontrado no Desktop.", "Arquivo não encontrado", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            txtVeriAntiparasitario.Text = "Verificar Anti-Parasitário:✅";
        }

        private void suplemento_Click(object sender, RoutedEventArgs e)
        {
            var relacaoWindow = new relacaoSuplemento();
            relacaoWindow.Show();
            lblGerarSuplmento.Text = "Gerar Suplemento:✅";
        }

        private void vrfcSuplemento_Click(object sender, RoutedEventArgs e)
        {
            // Obter o caminho do Desktop
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            // Combinar o caminho do Desktop com o nome do arquivo
            string filePath = Path.Combine(desktopPath, "Suplemento.xlsx");

            // Verificar se o arquivo existe
            if (File.Exists(filePath))
            {
                // Abrir o arquivo
                Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
            }
            else
            {
                MessageBox.Show("O arquivo Suplemento.xlsx não foi encontrado no Desktop.", "Arquivo não encontrado", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            txtVeriSuplemento.Text = "Verificar Suplemento:✅";
        }

        private void vermifugo_Click(object sender, RoutedEventArgs e)
        {
            var relacaoWindow = new relacaoVermifugo();
            relacaoWindow.Show();
            lblGerarVermifugo.Text = "Gerar Vermifugo:✅";
        }

        private void vrfcVermifugo_Click(object sender, RoutedEventArgs e)
        {

            // Obter o caminho do Desktop
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            // Combinar o caminho do Desktop com o nome do arquivo
            string filePath = Path.Combine(desktopPath, "Vermifugo.xlsx");

            // Verificar se o arquivo existe
            if (File.Exists(filePath))
            {
                // Abrir o arquivo
                Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
            }
            else
            {
                MessageBox.Show("O arquivo Vermifugo.xlsx não foi encontrado no Desktop.", "Arquivo não encontrado", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            txtVeriVermifugo.Text = "Verificar Vermifugo:✅";

        }

        private void racao_Click(object sender, RoutedEventArgs e)
        {
            var relacaoWindow = new relacaoRacao();
            relacaoWindow.Show();
            lblGerarRacao.Text = "Gerar Ração:✅";
        }
        private void vrfcRacao_Click(object sender, RoutedEventArgs e)
        {
            // Obter o caminho do Desktop
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            // Combinar o caminho do Desktop com o nome do arquivo
            string filePath = Path.Combine(desktopPath, "Racao.xlsx");

            // Verificar se o arquivo existe
            if (File.Exists(filePath))
            {
                // Abrir o arquivo
                Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
            }
            else
            {
                MessageBox.Show("O arquivo Racao.xlsx não foi encontrado no Desktop.", "Arquivo não encontrado", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            txtVeriRacao.Text = "Verificar Ração:✅";
        }

        private void welcome_Click(object sender, RoutedEventArgs e)
        {
            var relacaoWindows = new relacaoWelcome();
            relacaoWindows.Show();
            lblGerarWelcome.Text = "Gerar Welcome:✅";
        }

        private void vrfcWelcome_Click(object sender, RoutedEventArgs e)
        {
            // Obter o caminho do Desktop
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            // Combinar o caminho do Desktop com o nome do arquivo
            string filePath = Path.Combine(desktopPath, "Welcome.xlsx");

            // Verificar se o arquivo existe
            if (File.Exists(filePath))
            {
                // Abrir o arquivo
                Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
            }
            else
            {
                MessageBox.Show("O arquivo Welcome.xlsx não foi encontrado no Desktop.", "Arquivo não encontrado", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            txtVeriWelcome.Text = "Verificar Welcome:✅";
        }

        private void vacina_Click(object sender, RoutedEventArgs e)
        {
            var relacaoWindows = new relacaoVacina();
            relacaoWindows.Show();
            lblGerarVacina.Text = "Gerar Vacina:✅";
        }

        private void vrfcVacina_Click(object sender, RoutedEventArgs e)
        {
            // Obter o caminho do Desktop
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            // Combinar o caminho do Desktop com o nome do arquivo
            string filePath = Path.Combine(desktopPath, "Vacina.xlsx");

            // Verificar se o arquivo existe
            if (File.Exists(filePath))
            {
                // Abrir o arquivo
                Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
            }
            else
            {
                MessageBox.Show("O arquivo Vacina.xlsx não foi encontrado no Desktop.", "Arquivo não encontrado", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            txtVeriVacina.Text = "Verificar Vacina:✅";
        }

        private void milteforan_Click(object sender, RoutedEventArgs e)
        {
            var relacaoWindow = new relacaoMilteforan();
            relacaoWindow.Show();
            lblGerarMilteforan.Text = "Gerar Milteforan:✅";
        }

        private void vrfcMilteforan_Click(object sender, RoutedEventArgs e)
        {
            // Obter o caminho do Desktop
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            // Combinar o caminho do Desktop com o nome do arquivo
            string filePath = Path.Combine(desktopPath, "Milteforan.xlsx");

            // Verificar se o arquivo existe
            if (File.Exists(filePath))
            {
                // Abrir o arquivo
                Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
            }
            else
            {
                MessageBox.Show("O arquivo Milteforan.xlsx não foi encontrado no Desktop.", "Arquivo não encontrado", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            txtVeriMilteforan.Text = "Verificar Milteforan:✅";
        }

        private void relatorio_Click(object sender, RoutedEventArgs e)
        {
            var relacaoWindow = new relatorioRetorno();
            relacaoWindow.Show();
        }
     //-----------------------------------------------------------------------------------------------------------------------------------------------------------

    }
}
