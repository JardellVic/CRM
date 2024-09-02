using System.Windows;
using System.Windows.Input;

namespace CRM
{
    public partial class Index : Window
    {
        public Index()
        {
            InitializeComponent();
            this.KeyDown += new KeyEventHandler(EnterClick);
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
