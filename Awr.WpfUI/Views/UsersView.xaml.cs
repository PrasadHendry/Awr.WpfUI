using System.Windows;
using System.Windows.Controls;
using Awr.WpfUI.ViewModels;

namespace Awr.WpfUI.Views
{
    public partial class UsersView : UserControl
    {
        public UsersView()
        {
            InitializeComponent();
        }

        // Push PasswordBox -> ViewModel
        private void PasswordBox_PasswordChanged(object sender, RoutedEventArgs e)
        {
            if (this.DataContext is UsersViewModel vm && !vm.IsPasswordVisible)
            {
                vm.NewPassword = ((PasswordBox)sender).Password;
            }
        }

        // Push ViewModel -> PasswordBox (When switching back to hidden)
        private void EyeButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.DataContext is UsersViewModel vm)
            {
                // If we are about to hide (Visible -> Hidden), copy Text back to PasswordBox
                // Note: The Command runs first (toggling bool), so we check the NEW state
                if (!vm.IsPasswordVisible)
                {
                    pwdBox.Password = vm.NewPassword;
                }
            }
        }
    }
}