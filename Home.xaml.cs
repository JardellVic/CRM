﻿#region Diretivas
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using Microsoft.Win32;
using OfficeOpenXml;
using CRM.conexao.API;
using System.Windows.Threading;
using System.Text.Json;
using CRM.Ajuda;
using CRM.conexao;
#endregion

namespace CRM
{
    public class LineData
    {
        public string Numero { get; set; } = string.Empty;
        public string Nome { get; set; } = string.Empty;
        public List<string> Variaveis { get; set; } = [];
    }

    public partial class Home : Window
    {
        #region Properties

        string desktopPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Relações");
        public static Home Instance { get; private set; }
        public string TemplateIdSelecionado { get; set; } = string.Empty;
        public List<LineData> LinhasParaEnviar { get; private set; } = [];
        //-----------------------------------------------------------------------------------------------------------------------//
        private readonly APIManager apiManager;
        private readonly HttpClient client;
        private readonly Dictionary<string, string> templateTextMap;
        private readonly Dictionary<string, int> templateParamsMap;
        private readonly Dictionary<string, string> templateIdMap;
        private DispatcherTimer timer = new();
        private TimeSpan tempoRestante;
        private bool isDisparoPaused = false;
        private bool isDisparoRunning = false;
        private int currentIndex = 0;
        private conexaoCRM _conexaoCRM;
        private string _username;
        #endregion

        public Home(string username)
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            apiManager = new APIManager();
            client = new HttpClient();
            templateTextMap = [];
            templateParamsMap = [];
            templateIdMap = [];
            LoadTemplatesAsync();
            cmbTemplates.SelectionChanged += CmbTemplates_SelectionChanged;
            Instance = this;
            this.ResizeMode = ResizeMode.NoResize;
            SelectFileButton.IsEnabled = false;
            lblData.Content = DateTime.Now.ToString("dd/MM/yyyy");
            _username = username;
            lblUsuario.Content = username;
            _conexaoCRM = new conexaoCRM();
            AtualizarLabelsGerar();
            AtualizarLabelsVerificar();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Application.Current.Shutdown();
        }

        #region API PlanetFone

        private static void ShowError(string message)
        {
            MessageBox.Show(message, "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        private async void btnEnviarDisparo_Click(object sender, RoutedEventArgs e)
        {
            if (isDisparoPaused)
            {
                isDisparoPaused = false;
                isDisparoRunning = true;

                InicializarContagemRegressiva(Home.Instance.LinhasParaEnviar.Count - currentIndex);

                await ContinuarEnvioAsync();
            }
            else
            {
                if (!isDisparoRunning)
                {
                    isDisparoRunning = true;
                    InicializarContagemRegressiva(Home.Instance.LinhasParaEnviar.Count);
                    await ContinuarEnvioAsync();
                }
            }
        }

        private void btnPareDisparo_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show("Deseja finalizar o disparo?", "Confirmar", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result == MessageBoxResult.Yes)
            {
                isDisparoRunning = false;
                isDisparoPaused = false;

                timer?.Stop();

                txtBlockConsole.Inlines.Add(new Run("\nDisparo finalizado.") { Foreground = Brushes.Red });
            }
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
            try
            {
                statusContatos.Content = $"Quantidade de contatos: {quantidadeContatos}";

                double valorUtility = 0.008 * quantidadeContatos;
                statusUtility.Content = $"Valor Utility: ${valorUtility:F2}";

                double valorMarketing = 0.0625 * quantidadeContatos;
                statusMarketing.Content = $"Valor Marketing: ${valorMarketing:F2}";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ocorreu um erro ao atualizar a barra de status: {ex.Message}");
            }
        }

        private void CmbTemplates_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (cmbTemplates.SelectedItem != null)
                {
                    string selectedTemplate = cmbTemplates.SelectedItem?.ToString() ?? string.Empty;
                    if (templateTextMap.ContainsKey(selectedTemplate))
                    {
                        txtTemplate.Text = templateTextMap[selectedTemplate];
                        if (templateIdMap.ContainsKey(selectedTemplate))
                        {
                            TemplateIdSelecionado = templateIdMap[selectedTemplate];
                        }
                        SelectFileButton.IsEnabled = true;
                    }

                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void SelectFileButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new()
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    FilePathTextBox.Text = openFileDialog.FileName;

                    if (cmbTemplates.SelectedItem != null)
                    {
                        string selectedTemplate = cmbTemplates.SelectedItem?.ToString() ?? string.Empty;
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
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private static List<string> GetColumnNamesFromExcel(string filePath)
        {
            try
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
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return new List<string>();
            }

        }

        private static List<List<string>> GetRowDataFromExcel(string filePath)
        {
            var rowData = new List<List<string>>();

            try
            {
                using var package = new ExcelPackage(new FileInfo(filePath));
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
            try
            {
                var columnNames = GetColumnNamesFromExcel(FilePathTextBox.Text);
                var rowData = GetRowDataFromExcel(FilePathTextBox.Text);

                int quantidadeContatos = rowData.Count;
                AtualizarStatusBar(quantidadeContatos);

                MappingWindow mappingWindow = new(paramsCount, columnNames, rowData);
                mappingWindow.ShowDialog();

                if (!string.IsNullOrEmpty(mappingWindow.ColunaNumeroSelecionada) &&
                    !string.IsNullOrEmpty(mappingWindow.ColunaNomeSelecionada))
                {
                    ProcessRowData(mappingWindow.ColunaNumeroSelecionada, mappingWindow.ColunaNomeSelecionada, mappingWindow.ColunaVariaveisSelecionada, rowData);
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); };
        }

        private void ProcessRowData(string colunaNumero, string colunaNome, string variaveisColuna, List<List<string>> rowData)
        {
            try
            {
                var linhas = new List<LineData>();
                var columnNames = GetColumnNamesFromExcel(FilePathTextBox.Text);

                if (columnNames == null)
                {
                    ShowError("Não foi possível obter os nomes das colunas do Excel.");
                    return;
                }

                int numeroIndex = columnNames.IndexOf(colunaNumero);
                int nomeIndex = columnNames.IndexOf(colunaNome);

                if (numeroIndex == -1 || nomeIndex == -1)
                {
                    ShowError("As colunas especificadas não foram encontradas no Excel.");
                    return;
                }

                if (string.IsNullOrEmpty(variaveisColuna))
                {
                    MessageBox.Show("A coluna de variáveis está vazia.");
                    return;
                }

                List<int> variaveisIndices = variaveisColuna.Trim('[', ']').Split(',')
                    .Select(v => columnNames.IndexOf(v.Trim('"')))
                    .ToList();

                if (variaveisIndices.Any(i => i < 0 || i >= columnNames.Count))
                {
                    ShowError("Alguns índices de variáveis não correspondem às colunas do Excel.");
                    return;
                }

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
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private int GetParamsCountForTemplate(string templateName)
        {
            return templateParamsMap.ContainsKey(templateName) ? templateParamsMap[templateName] : 0;
        }

        private static int GetColumnCountFromExcel(string filePath)
        {
            try
            {
                using var package = new ExcelPackage(new FileInfo(filePath));
                var worksheet = package.Workbook.Worksheets[0];
                return worksheet.Dimension.End.Column;
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); return 0; }
        }

        private void InicializarContagemRegressiva(int quantidadeContatos)
        {
            try
            {
                double totalSegundos = 6 * quantidadeContatos;
                tempoRestante = TimeSpan.FromSeconds(totalSegundos);

                timer = new DispatcherTimer
                {
                    Interval = TimeSpan.FromSeconds(1)
                };
                #pragma warning disable CS8622
                timer.Tick += Timer_Tick;
                #pragma warning restore CS8622
                timer.Start();
            }
            catch (Exception e) { MessageBox.Show(e.Message); return; }
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            try
            {
                if (sender == null)
                {

                    return;
                }
                if (tempoRestante.TotalSeconds > 0)
                {
                    tempoRestante = tempoRestante.Subtract(TimeSpan.FromSeconds(1));
                    statusTempo.Content = $"Tempo Restante: {tempoRestante:hh\\:mm\\:ss}";
                }
                else
                {
                    timer.Stop();
                    statusTempo.Content = "Tempo Esgotado";
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private async Task ContinuarEnvioAsync()
        {
            int sucessoCount = 0;
            int erroCount = 0;
            int totalLinhas = Home.Instance.LinhasParaEnviar.Count;

            progressDisparo.IsIndeterminate = false;
            progressDisparo.Maximum = totalLinhas;
            progressDisparo.Value = currentIndex;

            txtBlockConsole.Inlines.Clear();
            txtBlockConsoleResponse.Inlines.Clear();
            txtBlockConsole.Inlines.Add(new Run($"Iniciando envio... ({currentIndex}/{totalLinhas})") { Foreground = Brushes.Yellow });

            scrollViewerConsole.ScrollToEnd();

            try
            {
                var linhasParaEnviar = Home.Instance.LinhasParaEnviar;

                for (int i = currentIndex; i < linhasParaEnviar.Count; i++)
                {
                    if (!isDisparoRunning)
                        break;

                    while (isDisparoPaused)
                    {
                        await Task.Delay(500); // Aguarda 0.5 segundo
                    }

                    bool resultado = await EnviarLinhaAsync(linhasParaEnviar[i]);
                    if (resultado)
                    {
                        sucessoCount++;
                    }
                    else
                    {
                        erroCount++;
                        txtBlockConsole.Inlines.Add(new Run($"\nErro: {linhasParaEnviar[i].Numero}") { Foreground = Brushes.Red });
                    }

                    currentIndex = i + 1;
                    progressDisparo.Value = currentIndex;
                    txtBlockConsole.Inlines.Add(new Run($"\nProgresso: {currentIndex}/{totalLinhas}") { Foreground = Brushes.Blue });

                    scrollViewerConsole.ScrollToEnd();
                    scrollViewerConsoleT.ScrollToEnd();

                    await Task.Delay(500); // Aguarda 0.5 segundo entre envios
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ocorreu um erro: {ex.Message}", "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                progressDisparo.IsIndeterminate = false;
                if (isDisparoRunning)
                {
                    MessageBoxResult result = MessageBox.Show(
                        $"Envios concluídos!\n\nSucessos: {sucessoCount}\nErros: {erroCount}",
                        "Resultado do Envio", MessageBoxButton.OK, MessageBoxImage.Information
                    );

                    if (result == MessageBoxResult.OK)
                    {
                        
                        txtBlockConsole.Inlines.Clear();
                        txtBlockConsoleResponse.Inlines.Clear();
                        isDisparoRunning = false;
                        isDisparoPaused = false;
                    }
                }           
                currentIndex = 0;
            }

        }

        private async Task<bool> EnviarLinhaAsync(LineData linha)
        {
            try
            {
                var (optinResult, responseBody) = await apiManager.OptinNumeroAsync(linha.Numero);
                (var sendResult, string responseBodyT) = await apiManager.EnviarLinhaAsync(TemplateIdSelecionado, linha.Numero, linha.Nome, linha.Variaveis);

                string formattedResponseBody = FormatJson(responseBody);
                string formattedResponseBodyT = FormatJson(responseBodyT);

                // Display formatted response
                if (optinResult)
                {
                    txtBlockConsoleResponse.Inlines.Add(new Run($"\nOptIn feito com sucesso:") { Foreground = Brushes.Yellow });
                    txtBlockConsoleResponse.Inlines.Add(new Run($"\n{formattedResponseBody}") { Foreground = Brushes.Green });
                    txtBlockConsoleResponse.Inlines.Add(new Run($"\nMensagem enviada com sucesso") { Foreground = Brushes.Yellow });
                    txtBlockConsoleResponse.Inlines.Add(new Run($"\n{formattedResponseBodyT}") { Foreground = Brushes.Green });
                }
                else
                {
                    txtBlockConsoleResponse.Inlines.Add(new Run($"\nErro no OptIn:") { Foreground = Brushes.Red });
                    txtBlockConsoleResponse.Inlines.Add(new Run($"\n{formattedResponseBody}") { Foreground = Brushes.Red });
                    txtBlockConsoleResponse.Inlines.Add(new Run($"\n{formattedResponseBodyT}") { Foreground = Brushes.Red });
                    txtBlockConsole.Inlines.Add(new Run("Erro ao fazer optin") { Foreground = Brushes.Red });
                    return false;
                }

                if (sendResult)
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

        private string FormatJson(string json)
        {
            try
            {
                var jsonObject = JsonSerializer.Deserialize<JsonElement>(json);

                if (jsonObject.TryGetProperty("data", out var dataProperty))
                {
                    if (dataProperty.TryGetProperty("msg", out var msgProperty))
                    {
                        var status = jsonObject.GetProperty("status").GetString();
                        var message = msgProperty.GetString();
                        return $"Status: {status}\nMensagem: {message}";
                    }
                    else if (dataProperty.TryGetProperty("messageId", out var messageIdProperty) &&
                             dataProperty.TryGetProperty("status", out var statusProperty))
                    {
                        // Formato com 'messageId' e 'status'
                        var status = jsonObject.GetProperty("status").GetString();
                        var messageId = messageIdProperty.GetString();
                        var submittedStatus = statusProperty.GetString();
                        return $"Status: {status}\nMessageId: {messageId}\nStatus: {submittedStatus}";
                    }
                }

                // Caso não consiga identificar o formato, retorna o JSON bruto
                return json;
            }
            catch
            {
                // Retorna o JSON bruto se a formatação falhar
                return json;
            }
        }

        private async Task<bool> OptinNumeroAsync(string numero)
        {
            try
            {
                var (result, responseBody) = await apiManager.OptinNumeroAsync(numero);

                return result;
            }
            catch (Exception)
            {
                return false;
            }
        }

        #endregion

        #region Listas 
        private void BancoMenuItem_Click(object sender, RoutedEventArgs e)
        {
            AtualizarBanco atualizarBancoPage = new();
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
            string filePath = Path.Combine(desktopPath, "Antiparasitario.xlsx");

            if (File.Exists(filePath))
            {
                Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
                _conexaoCRM.AtualizarExecucaoVerificar("antiparasitario");
                txtVeriAntiparasitario.Text = "Verificar Anti-Parasitário:✅";
            }
            else
            {
                MessageBox.Show("O arquivo Antiparasitario.xlsx não foi encontrado no Desktop.", "Arquivo não encontrado", MessageBoxButton.OK, MessageBoxImage.Warning);
            }

        }

        private void HelpAntiparasitario_Click(object sender, RoutedEventArgs e)
        {
            HelpAntiparasitario helpantiparasitario = new HelpAntiparasitario();
            helpantiparasitario.ShowDialog();
        }

        private void suplemento_Click(object sender, RoutedEventArgs e)
        {
            var relacaoWindow = new relacaoSuplemento();
            relacaoWindow.Show();
            lblGerarSuplmento.Text = "Gerar Suplemento:✅";
        }

        private void vrfcSuplemento_Click(object sender, RoutedEventArgs e)
        {
            string filePath = Path.Combine(desktopPath, "Suplemento.xlsx");

            if (File.Exists(filePath))
            {
                Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
                _conexaoCRM.AtualizarExecucaoVerificar("suplemento");
                txtVeriSuplemento.Text = "Verificar Suplemento:✅";
            }
            else
            {
                MessageBox.Show("O arquivo Suplemento.xlsx não foi encontrado no Desktop.", "Arquivo não encontrado", MessageBoxButton.OK, MessageBoxImage.Warning);
            }   
        }

        private void HelpSuplemento_Click(object sender, RoutedEventArgs e)
        {
            helpSuplemento helpsuplemento = new helpSuplemento();
            helpsuplemento.ShowDialog();
        }

        private void vermifugo_Click(object sender, RoutedEventArgs e)
        {
            var relacaoWindow = new relacaoVermifugo();
            relacaoWindow.Show();
            lblGerarVermifugo.Text = "Gerar Vermifugo:✅";
        }

        private void vrfcVermifugo_Click(object sender, RoutedEventArgs e)
        {
            string filePath = Path.Combine(desktopPath, "Vermifugo.xlsx");

            if (File.Exists(filePath))
            {
                Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
                _conexaoCRM.AtualizarExecucaoVerificar("vermifugo");
                txtVeriVermifugo.Text = "Verificar Vermifugo:✅";
            }
            else
            {
                MessageBox.Show("O arquivo Vermifugo.xlsx não foi encontrado no Desktop.", "Arquivo não encontrado", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void racao_Click(object sender, RoutedEventArgs e)
        {
            var relacaoWindow = new relacaoRacao();
            relacaoWindow.Show();
            lblGerarRacao.Text = "Gerar Ração:✅";
        }

        private void vrfcRacao_Click(object sender, RoutedEventArgs e)
        {
            string filePath = Path.Combine(desktopPath, "Racao.xlsx");

            if (File.Exists(filePath))
            {
                Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
                _conexaoCRM.AtualizarExecucaoVerificar("racao");
                txtVeriRacao.Text = "Verificar Ração:✅";
            }
            else
            {
                MessageBox.Show("O arquivo Racao.xlsx não foi encontrado no Desktop.", "Arquivo não encontrado", MessageBoxButton.OK, MessageBoxImage.Warning);
            }     
        }

        private void welcome_Click(object sender, RoutedEventArgs e)
        {
            var relacaoWindows = new relacaoWelcome();
            relacaoWindows.Show();
            lblGerarWelcome.Text = "Gerar Welcome:✅";
        }

        private void vrfcWelcome_Click(object sender, RoutedEventArgs e)
        {
            string filePath = Path.Combine(desktopPath, "Welcome.xlsx");

            if (File.Exists(filePath))
            {
                Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
                _conexaoCRM.AtualizarExecucaoVerificar("welcome");
                txtVeriWelcome.Text = "Verificar Welcome:✅";
            }
            else
            {
                MessageBox.Show("O arquivo Welcome.xlsx não foi encontrado no Desktop.", "Arquivo não encontrado", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void vacina_Click(object sender, RoutedEventArgs e)
        {
            var relacaoWindows = new relacaoVacina();
            relacaoWindows.Show();
            lblGerarVacina.Text = "Gerar Vacina:✅";
        }

        private void vrfcVacina_Click(object sender, RoutedEventArgs e)
        {
            string filePath = Path.Combine(desktopPath, "Vacina.xlsx");

            if (File.Exists(filePath))
            {
                Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
                _conexaoCRM.AtualizarExecucaoVerificar("vacina");
                txtVeriVacina.Text = "Verificar Vacina:✅";
            }
            else
            {
                MessageBox.Show("O arquivo Vacina.xlsx não foi encontrado no Desktop.", "Arquivo não encontrado", MessageBoxButton.OK, MessageBoxImage.Warning);
            }           
        }

        private void milteforan_Click(object sender, RoutedEventArgs e)
        {
            var relacaoWindow = new relacaoMilteforan();
            relacaoWindow.Show();
            lblGerarMilteforan.Text = "Gerar Milteforan:✅";
        }

        private void vrfcMilteforan_Click(object sender, RoutedEventArgs e)
        {
            string filePath = Path.Combine(desktopPath, "Milteforan.xlsx");

            if (File.Exists(filePath))
            {
                Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
                _conexaoCRM.AtualizarExecucaoVerificar("milteforan");
                txtVeriMilteforan.Text = "Verificar Milteforan:✅";
            }
            else
            {
                MessageBox.Show("O arquivo Milteforan.xlsx não foi encontrado no Desktop.", "Arquivo não encontrado", MessageBoxButton.OK, MessageBoxImage.Warning);
            }   
        }

        private void relatorio_Click(object sender, RoutedEventArgs e)
        {
            var relacaoWindow = new relatorioRetorno();
            relacaoWindow.Show();
        }

        private void clientesPorProduto_Click(object sender, RoutedEventArgs e)
        {
            var relacaoWindow = new relacaoCpP();
            relacaoWindow.Show();
        }

        private void AtualizarLabelsGerar()
        {
            try
            {
                var controle = _conexaoCRM.ObterControleExecucaoGerar();

                if (controle != null)
                {
                    lblGerarAntiparasitario.Text = controle.Antiparasitario ? "Gerar Anti-Parasitário:✅" : "Gerar Anti-Parasitário:";
                    lblGerarSuplmento.Text = controle.Suplemento ? "Gerar Suplemento:✅" : "Gerar Suplemento:";
                    lblGerarVermifugo.Text = controle.Vermifugo ? "Gerar Vermífugo:✅" : "Gerar Vermífugo:";
                    lblGerarRacao.Text = controle.Racao ? "Gerar Ração:✅" : "Gerar Ração:";
                    lblGerarWelcome.Text = controle.Welcome ? "Gerar Welcome:✅" : "Gerar Welcome:";
                    lblGerarVacina.Text = controle.Vacina ? "Gerar Vacina:✅" : "Gerar Vacina:";
                    lblGerarMilteforan.Text = controle.Milteforan ? "Gerar Milteforan:✅" : "Gerar Milteforan:";
                }
                else
                {
                    MessageBox.Show("REGISTRO DE EXECUÇÃO NÃO CRIADO");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao atualizar labels: {ex.Message}");
            }
        }

        private void AtualizarLabelsVerificar()
        {
            try
            {
                var controle = _conexaoCRM.ObterControleExecucaoVerificar();

                if (controle != null)
                {
                    txtVeriAntiparasitario.Text = controle.Antiparasitario ? "Verificar Anti-Parasitário:✅" : "Verificar Anti-Parasitário:";
                    txtVeriSuplemento.Text = controle.Suplemento ? "Verificar Suplemento:✅" : "Verificar Suplemento:";
                    txtVeriVermifugo.Text = controle.Vermifugo ? "Verificar Vermífugo:✅" : "Verificar Vermífugo:";
                    txtVeriRacao.Text = controle.Racao ? "Verificar Ração:✅" : "Verificar Ração:";
                    txtVeriWelcome.Text = controle.Welcome ? "Verificar Welcome:✅" : "Verificar Welcome:";
                    txtVeriVacina.Text = controle.Vacina ? "Verificar Vacina:✅" : "Verificar Vacina:";
                    txtVeriMilteforan.Text = controle.Milteforan ? "Verificar Milteforan:✅" : "Verificar Milteforan:";
                }
                else
                {
                    MessageBox.Show("REGISTRO DE EXECUÇÃO NÃO CRIADO");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao atualizar labels: {ex.Message}");
            }
        }
    }
    #endregion
}