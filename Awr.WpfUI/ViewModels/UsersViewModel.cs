using System;
using System.Collections.ObjectModel;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using Awr.Core.DTOs;
using Awr.WpfUI.MvvmCore;
using Awr.WpfUI.Services.Implementation;
using Awr.WpfUI.Services.Interfaces;

namespace Awr.WpfUI.ViewModels
{
    public class UsersViewModel : BaseViewModel
    {
        private readonly IAuthenticationService _authService;

        public ObservableCollection<UserRoleDto> Roles { get; } = new ObservableCollection<UserRoleDto>();

        private UserRoleDto _selectedRole;
        public UserRoleDto SelectedRole
        {
            get => _selectedRole;
            set 
            { 
                if (SetProperty(ref _selectedRole, value))
                {
                    NewPassword = ""; // Clear on selection change
                    IsPasswordVisible = false; // Reset visibility
                }
            }
        }

        private string _newPassword;
        public string NewPassword { get => _newPassword; set => SetProperty(ref _newPassword, value); }

        private bool _isPasswordVisible;
        public bool IsPasswordVisible { get => _isPasswordVisible; set => SetProperty(ref _isPasswordVisible, value); }

        private bool _isBusy;
        public bool IsBusy { get => _isBusy; set { SetProperty(ref _isBusy, value); CommandManager.InvalidateRequerySuggested(); } }

        public ICommand RefreshCommand { get; }
        public ICommand UpdatePasswordCommand { get; }
        public ICommand TogglePasswordVisibilityCommand { get; }

        public UsersViewModel()
        {
            _authService = new AuthenticationService();
            
            RefreshCommand = new RelayCommand(async _ => await LoadRolesAsync());
            
            UpdatePasswordCommand = new RelayCommand(async _ => await UpdatePasswordAsync(), 
                _ => SelectedRole != null && !string.IsNullOrWhiteSpace(NewPassword) && !IsBusy);

            TogglePasswordVisibilityCommand = new RelayCommand(_ => IsPasswordVisible = !IsPasswordVisible);

            _ = LoadRolesAsync();
        }

        private async Task LoadRolesAsync()
        {
            IsBusy = true;
            try
            {
                var list = await _authService.GetUserRolesAsync();
                Roles.Clear();
                foreach (var r in list) Roles.Add(r);
            }
            catch (Exception ex) { MessageBox.Show("Error loading roles: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error); }
            finally { IsBusy = false; }
        }

        private async Task UpdatePasswordAsync()
        {
            if (SelectedRole == null) return;
            if (MessageBox.Show($"Update password for role '{SelectedRole.RoleName}'?", "Confirm", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.No) return;
            IsBusy = true;
            try
            {
                await _authService.UpdatePasswordAsync(SelectedRole.RoleName, NewPassword);
                MessageBox.Show("Password updated.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                NewPassword = "";
            }
            catch (Exception ex) { MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error); }
            finally { IsBusy = false; }
        }
    }
}