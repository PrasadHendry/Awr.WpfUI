using System.Windows;
using Awr.WpfUI.ViewModels;

namespace Awr.WpfUI.Views
{
    public partial class MainWindow : Window
    {
        // Constructor with 2 arguments
        public MainWindow(string role, string pcUser)
        {
            InitializeComponent();
            this.DataContext = new MainViewModel(role, pcUser);
        }
    }
}