using System.Diagnostics;
using System.Reflection;
using System.Windows;
using System.Windows.Input;
using CRM.conexao;

namespace CRM
{
    public partial class Index : Window
    {
        private readonly conexaoCRM conexao;

        public Index()
        {
            InitializeComponent();
            conexao = new conexaoCRM();

            string localVersion = GetAssemblyVersion();
            string dbVersion = string.Empty;
            txtPass.Focus();

            try
            {
                dbVersion = conexao.GetDatabaseVersion();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erro de Conexão", MessageBoxButton.OK, MessageBoxImage.Error);
                Application.Current.Shutdown();
                return;
            }

            if (dbVersion == localVersion)
            {
                this.KeyDown += new KeyEventHandler(EnterClick);
                txtLogin.Focus();
            }
            else
            {
                MessageBox.Show($"Seu sistema está desatualizado.\n Versão atual: {localVersion}\n Última versão: {dbVersion}\n", "Atualização Necessária", MessageBoxButton.OK, MessageBoxImage.Warning);
                Process.Start("AtualizadorCRM.exe");
                Application.Current.Shutdown();
            }
        }

        private string GetAssemblyVersion()
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            Version version = assembly.GetName().Version;
            return version.ToString();
        }

        private void EnterClick(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                BtnLogin_Click(sender, new RoutedEventArgs());
            }
        }

        private void BtnLogin_Click(object sender, RoutedEventArgs e)
        {
            string login = txtLogin.Text;
            string senha = txtPass.Password;

            string username = string.Empty;
            try
            {
                username = conexao.VerificarCredenciais(login, senha);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erro de Conexão", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (!string.IsNullOrEmpty(username))
            {
                try
                {
                    conexao.InserirControleExecucao(username);
                    conexao.InserirControleExecucaoVerificar(username);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Erro ao registrar controle de execução", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Redireciona para a próxima tela
                Home homeWindow = new Home(username);
                homeWindow.Show();
                this.Close();
            }
            else
            {
                MessageBox.Show("Login ou senha incorretos. Tente novamente.", "Erro de Autenticação", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

    }
}
