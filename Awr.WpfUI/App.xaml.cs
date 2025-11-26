using System.Windows;
using Awr.WpfUI.Views;
using Awr.WpfUI.ViewModels;

namespace Awr.WpfUI
{
    public partial class App : Application
    {
        private void Application_Startup(object sender, StartupEventArgs e)
        {
            // 1. Show Login Dialog
            LoginView loginView = new LoginView();
            bool? result = loginView.ShowDialog();

            // 2. Check Result
            if (result == true)
            {
                var loginVm = loginView.DataContext as LoginViewModel;

                // 3. Create Main Window
                MainWindow main = new MainWindow(loginVm.SelectedRole.RoleName, System.Environment.UserName);

                // CRITICAL FIX: Set the new Main Window and Reset Shutdown Mode
                this.MainWindow = main;
                this.ShutdownMode = ShutdownMode.OnMainWindowClose;

                main.Show();
            }
            else
            {
                // Login failed or cancelled
                this.Shutdown();
            }
        }
    }
}