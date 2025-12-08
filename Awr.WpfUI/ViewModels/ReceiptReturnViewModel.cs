using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using Awr.Core.DTOs;
using Awr.WpfUI.MvvmCore;
using Awr.WpfUI.ViewModels.Shared;

namespace Awr.WpfUI.ViewModels
{
    public class ReceiptReturnViewModel : WorkQueueViewModel
    {
        private string _voidReason;
        public string VoidReason { get => _voidReason; set => SetProperty(ref _voidReason, value); }

        // --- Error State ---
        private bool _isVoidError;
        public bool IsVoidError { get => _isVoidError; set => SetProperty(ref _isVoidError, value); }

        public ICommand PrintCommand { get; }
        public ICommand VoidCommand { get; }

        public ReceiptReturnViewModel(string username) : base(username)
        {
            PrintCommand = new RelayCommand(async _ => await PrintAsync(), _ => SelectedItem != null);
            VoidCommand = new RelayCommand(async _ => await VoidAsync(), _ => SelectedItem != null);
        }

        protected override async Task<List<AwrItemQueueDto>> FetchDataInternalAsync() => await Service.GetReceiptQueueAsync(Username);

        protected override void OnSelectionChanged() { VoidReason = string.Empty; IsVoidError = false; CommandManager.InvalidateRequerySuggested(); }

        private async Task PrintAsync()
        {
            if (SelectedItem == null) return;

            // Updated MessageBox with Asterisk/Information Icon
            if (MessageBox.Show($"Print Request {SelectedItem.RequestNo}?", "Confirm",
                MessageBoxButton.YesNo, MessageBoxImage.Asterisk) == MessageBoxResult.No)
                return;

            IsBusy = true;
            try
            {
                await Service.PrintAndReceiveItemAsync(SelectedItem.ItemId, Username);
                MessageBox.Show("Printed.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                await LoadDataAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsBusy = false;
            }
        }

        private async Task VoidAsync()
        {
            if (SelectedItem == null) return;
            IsVoidError = false;
            if (string.IsNullOrWhiteSpace(VoidReason))
            {
                IsVoidError = true;
                MessageBox.Show("Enter Void Reason.", "Validation", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            if (MessageBox.Show($"Void Request {SelectedItem.RequestNo}?", "Confirm", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.No) return;
            IsBusy = true;
            try { await Service.VoidItemAsync(SelectedItem.ItemId, Username, VoidReason); MessageBox.Show("Voided.", "Success", MessageBoxButton.OK, MessageBoxImage.Information); await LoadDataAsync(); }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Exclamation); }
            finally { IsBusy = false; }
        }
    }
}