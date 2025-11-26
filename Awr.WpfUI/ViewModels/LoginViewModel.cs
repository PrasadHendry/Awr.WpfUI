using System;
using System.Collections.ObjectModel;
using System.Threading.Tasks;
using System.Windows.Input;
using Awr.Core.DTOs;
using Awr.WpfUI.MvvmCore;
using Awr.WpfUI.Services.Implementation;
using Awr.WpfUI.Services.Interfaces;

namespace Awr.WpfUI.ViewModels
{
    public class LoginViewModel : BaseViewModel
    {
        private readonly IAuthenticationService _authService;

        // --- State Properties ---
        private bool _isBusy;
        public bool IsBusy { get => _isBusy; set => SetProperty(ref _isBusy, value); }

        private string _errorMessage;
        public string ErrorMessage { get => _errorMessage; set => SetProperty(ref _errorMessage, value); }

        private ObservableCollection<UserRoleDto> _roles;
        public ObservableCollection<UserRoleDto> Roles { get => _roles; set => SetProperty(ref _roles, value); }

        private UserRoleDto _selectedRole;
        public UserRoleDto SelectedRole { get => _selectedRole; set => SetProperty(ref _selectedRole, value); }

        private string _password;
        public string Password { get => _password; set => SetProperty(ref _password, value); }

        // --- Events ---
        public Action<string, string> OnLoginSuccess; // (Username, Role)

        // --- Commands ---
        public ICommand LoadRolesCommand { get; }
        public ICommand LoginCommand { get; }

        public LoginViewModel()
        {
            // Ideally use DI here, but manual injection for simplicity per Dev Plan
            _authService = new AuthenticationService();
            Roles = new ObservableCollection<UserRoleDto>();

            LoadRolesCommand = new RelayCommand(async _ => await LoadRolesAsync());
            LoginCommand = new RelayCommand(async _ => await AttemptLoginAsync(), CanLogin);
        }

        private bool CanLogin(object obj)
        {
            return !IsBusy && SelectedRole != null && !string.IsNullOrWhiteSpace(Password);
        }

        private async Task LoadRolesAsync()
        {
            IsBusy = true;
            ErrorMessage = "Connecting to database...";
            try
            {
                var roles = await _authService.GetUserRolesAsync();
                Roles.Clear();
                foreach (var r in roles) Roles.Add(r);

                if (Roles.Count > 0) SelectedRole = Roles[0];
                ErrorMessage = string.Empty;
            }
            catch (Exception ex)
            {
                ErrorMessage = $"Connection Failed: {ex.Message}";
            }
            finally
            {
                IsBusy = false;
            }
        }

        private async Task AttemptLoginAsync()
        {
            if (SelectedRole == null) return;

            IsBusy = true;
            ErrorMessage = "Verifying credentials...";

            try
            {
                var result = await _authService.ValidateUserAsync(SelectedRole.RoleName, Password);

                if (result.IsSuccess)
                {
                    OnLoginSuccess?.Invoke(SelectedRole.RoleName, result.Role);
                }
                else
                {
                    ErrorMessage = "Invalid Password.";
                }
            }
            catch (Exception ex)
            {
                ErrorMessage = $"Login Error: {ex.Message}";
            }
            finally
            {
                IsBusy = false;
                // Force command manager to re-evaluate buttons (enable/disable)
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }
}