using System.Reflection;
using Awr.WpfUI.MvvmCore;

namespace Awr.WpfUI.ViewModels
{
    public class AboutViewModel : BaseViewModel
    {
        public string AppName => "AWR Issuance System (AWR-IS)";
        public string AppVersion { get; }
        public string Copyright => "© 2025 Sigma Laboratories Pvt. Ltd.";
        public string SupportContact => "Contact IT Support: Extension 231";

        public AboutViewModel()
        {
            // Get Version from Assembly
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            AppVersion = $"Version: {version.Major}.{version.Minor}.{version.Build}.{version.Revision} (Beta 1)";
        }
    }
}