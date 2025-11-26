using System.Windows;
using System.Windows.Controls;
using Awr.WpfUI.ViewModels;

namespace Awr.WpfUI.Views
{
    public partial class LoginView : Window
    {
        public LoginView()
        {
            InitializeComponent();
            var vm = new LoginViewModel();
            this.DataContext = vm;

            // Wire up the close action
            vm.OnLoginSuccess += (username, role) =>
            {
                // We will implement MainWindow opening later. 
                // For now, this DialogResult=true signals success to App.xaml.
                this.DialogResult = true;
                this.Close();
            };

            // Trigger loading roles
            this.Loaded += (s, e) => vm.LoadRolesCommand.Execute(null);
        }

        private void PasswordBox_PasswordChanged(object sender, RoutedEventArgs e)
        {
            if (this.DataContext is LoginViewModel vm)
            {
                vm.Password = ((PasswordBox)sender).Password;
            }
        }
    }
}