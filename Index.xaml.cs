using System.Diagnostics;
using System.Reflection;
using System.Windows;
using System.Windows.Input;
using Npgsql;

namespace CRM
{
    public partial class Index : Window
    {
        private const string connectionString = "Host=172.16.1.103;Port=5432;Username=jvsilva;Password=1011;Database=crm"; 
        public Index()
        {
            InitializeComponent();

            string localVersion = GetAssemblyVersion();

            string dbVersion = GetDatabaseVersion();

            if (dbVersion == localVersion)
            {
                this.KeyDown += new KeyEventHandler(EnterClick);
                txtPass.Focus();
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

        private string GetDatabaseVersion()
        {
            string dbVersion = string.Empty;

            try
            {
                using (var connection = new NpgsqlConnection(connectionString))
                {
                    connection.Open();

                    string query = "SELECT version FROM version LIMIT 1";
                    using (var command = new NpgsqlCommand(query, connection))
                    {
                        var result = command.ExecuteScalar();
                        dbVersion = result?.ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao conectar ao banco de dados: {ex.Message}", "Erro de Conexão", MessageBoxButton.OK, MessageBoxImage.Error);
                Application.Current.Shutdown();
            }

            return dbVersion;
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
            if (txtPass.Password == "489524")
            {
                Home homeWindow = new Home();
                homeWindow.Show();
                this.Close();
            }
            else
            {
                MessageBox.Show("Senha incorreta. Tente novamente.");
            }
        }
    }
}
