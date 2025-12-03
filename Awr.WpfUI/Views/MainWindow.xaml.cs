using System.ComponentModel;
using System.Windows;
using Awr.WpfUI.ViewModels;

namespace Awr.WpfUI.Views
{
    public partial class MainWindow : Window
    {
        private bool _isSigningOut = false;

        public MainWindow(string role, string pcUser)
        {
            InitializeComponent();

            var vm = new MainViewModel(role, pcUser);
            this.DataContext = vm;

            // Hook up SignOut event from ViewModel
            vm.SigningOut += (s, e) => _isSigningOut = true;

            // Hook up Window Closing event
            this.Closing += MainWindow_Closing;
        }

        private void MainWindow_Closing(object sender, CancelEventArgs e)
        {
            // If the user clicked "Sign Out", skip the confirmation
            if (_isSigningOut) return;

            // Otherwise (User clicked X), confirm exit
            var result = MessageBox.Show("Are you sure you want to exit the application?",
                                         "Confirm Exit",
                                         MessageBoxButton.YesNo,
                                         MessageBoxImage.Question);

            if (result == MessageBoxResult.No)
            {
                e.Cancel = true; // Stop closing
            }
        }
    }
}