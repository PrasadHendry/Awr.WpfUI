using System;
using System.Collections.ObjectModel;
using System.Windows.Input;
using Awr.WpfUI.MvvmCore;

namespace Awr.WpfUI.ViewModels
{
    public class MainViewModel : BaseViewModel
    {
        private string _loggedInUser;
        public string LoggedInUser { get => _loggedInUser; set => SetProperty(ref _loggedInUser, value); }

        private string _footerTime;
        public string FooterTime { get => _footerTime; set => SetProperty(ref _footerTime, value); }

        public ObservableCollection<TabItemViewModel> Tabs { get; set; } = new ObservableCollection<TabItemViewModel>();

        private TabItemViewModel _selectedTab;
        public TabItemViewModel SelectedTab { get => _selectedTab; set => SetProperty(ref _selectedTab, value); }

        public ICommand SignOutCommand { get; }
        private readonly string _pcUsername;

        public MainViewModel(string roleName, string pcUsername)
        {
            _pcUsername = pcUsername;
            LoggedInUser = $"{roleName} | {pcUsername}";

            var timer = new System.Windows.Threading.DispatcherTimer();
            timer.Interval = TimeSpan.FromSeconds(1);
            timer.Tick += (s, e) => FooterTime = DateTime.Now.ToString("dd-MM-yyyy hh:mm tt");
            timer.Start();

            SignOutCommand = new RelayCommand(_ => OnSignOut());

            ConfigureTabs(roleName);
        }

        private void ConfigureTabs(string role)
        {
            Tabs.Add(new TabItemViewModel("Audit Trail", new AuditTrailViewModel(_pcUsername)));

            switch (role)
            {
                case "Requester":
                    // UPDATE: Use NewRequestViewModel
                    Tabs.Add(new TabItemViewModel("New Request", new NewRequestViewModel(_pcUsername)));
                    Tabs.Add(new TabItemViewModel("Receipt & Print", new ReceiptReturnViewModel(_pcUsername)));
                    break;

                case "QA":
                    Tabs.Add(new TabItemViewModel("Approval Queue", new IssuanceQueueViewModel(_pcUsername)));
                    break;

                case "Admin":
                    // UPDATE: Use NewRequestViewModel
                    Tabs.Add(new TabItemViewModel("New Request", new NewRequestViewModel(_pcUsername)));
                    Tabs.Add(new TabItemViewModel("Approval Queue", new IssuanceQueueViewModel(_pcUsername)));
                    Tabs.Add(new TabItemViewModel("Receipt & Print", new ReceiptReturnViewModel(_pcUsername)));
                    Tabs.Add(new TabItemViewModel("Users", new PlaceholderViewModel()));
                    break;
            }

            if (Tabs.Count > 0) SelectedTab = Tabs[0];
        }

        private void OnSignOut()
        {
            System.Windows.Forms.Application.Restart();
            System.Windows.Application.Current.Shutdown();
        }
    }

    public class TabItemViewModel
    {
        public string Header { get; }
        public BaseViewModel Content { get; }
        public TabItemViewModel(string header, BaseViewModel content) { Header = header; Content = content; }
    }

    public class PlaceholderViewModel : BaseViewModel { }
}