using System.Windows;

namespace CRM
{
    /// <summary>
    /// Interaction logic for Index.xaml
    /// </summary>
    public partial class Index : Window
    {
        public Index()
        {
            InitializeComponent();
        }

        private void BtnLogin_Click(object sender, RoutedEventArgs e)
        {
            // Verifica se a senha inserida está correta
            if (txtPass.Password == "489524")
            {

                // Cria e abre a nova janela Home
                Home homeWindow = new Home();
                homeWindow.Show();

                // Fecha a janela atual
                this.Close();
            }
            else
            {
                MessageBox.Show("Senha incorreta. Tente novamente.");
            }
        }
    }
}
