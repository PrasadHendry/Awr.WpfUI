using Awr.WpfUI.MvvmCore;
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;

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

        // --- NEW: Event to notify View regarding SignOut vs Exit ---
        public event EventHandler SigningOut;

        public MainViewModel(string roleName, string pcUsername)
        {
            _pcUsername = pcUsername;
            LoggedInUser = $"{roleName} | {pcUsername}";

            var timer = new System.Windows.Threading.DispatcherTimer();
            timer.Interval = TimeSpan.FromSeconds(1);
            timer.Tick += (s, e) => FooterTime = DateTime.Now.ToString("dd-MM-yyyy HH:mm tt");
            timer.Start();

            SignOutCommand = new RelayCommand(_ => OnSignOut());

            ConfigureTabs(roleName);
        }

        private void ConfigureTabs(string role)
        {
            // 1. Add Audit Trail (Always First)
            Tabs.Add(new TabItemViewModel("Audit Trail", new AuditTrailViewModel(_pcUsername)));

            // 2. Add Role-Specific Tabs
            switch (role)
            {
                case "Requester":
                    Tabs.Add(new TabItemViewModel("New Request", new NewRequestViewModel(_pcUsername)));
                    Tabs.Add(new TabItemViewModel("Receipt & Print", new ReceiptReturnViewModel(_pcUsername)));
                    break;

                case "QA":
                    Tabs.Add(new TabItemViewModel("Approval Queue", new IssuanceQueueViewModel(_pcUsername)));
                    break;

                case "Admin":
                    Tabs.Add(new TabItemViewModel("New Request", new NewRequestViewModel(_pcUsername)));
                    Tabs.Add(new TabItemViewModel("Approval Queue", new IssuanceQueueViewModel(_pcUsername)));
                    Tabs.Add(new TabItemViewModel("Receipt & Print", new ReceiptReturnViewModel(_pcUsername)));
                    Tabs.Add(new TabItemViewModel("Users", new UsersViewModel()));
                    break;
            }

            // 3. Smart Tab Selection
            if (Tabs.Count > 0)
            {
                TabItemViewModel targetTab = Tabs[0];

                if (role == "Requester")
                {
                    var newReq = Tabs.FirstOrDefault(t => t.Header == "New Request");
                    if (newReq != null) targetTab = newReq;
                }
                else if (role == "QA")
                {
                    var approval = Tabs.FirstOrDefault(t => t.Header == "Approval Queue");
                    if (approval != null) targetTab = approval;
                }
                else if (role == "Admin")
                {
                    // FIX: Default to "Users" for Admin
                    var usersTab = Tabs.FirstOrDefault(t => t.Header == "Users");
                    if (usersTab != null) targetTab = usersTab;
                }

                SelectedTab = targetTab;
            }
        }

        private void OnSignOut()
        {
            // 1. Confirm Intent
            if (System.Windows.MessageBox.Show("Are you sure you want to Sign Out?", "Confirm Sign Out", 
                System.Windows.MessageBoxButton.YesNo, System.Windows.MessageBoxImage.Question) == System.Windows.MessageBoxResult.No)
            {
                return;
            }

            // 2. Raise Event (Tells View to disable the closing confirmation)
            SigningOut?.Invoke(this, EventArgs.Empty);

            // 3. Restart Application
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