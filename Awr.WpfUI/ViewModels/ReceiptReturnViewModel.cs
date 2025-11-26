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
        // --- State ---
        private string _voidReason;
        public string VoidReason
        {
            get => _voidReason;
            set => SetProperty(ref _voidReason, value);
        }

        // --- Commands ---
        public ICommand PrintCommand { get; }
        public ICommand VoidCommand { get; }

        // --- Constructor ---
        public ReceiptReturnViewModel(string username) : base(username)
        {
            // Commands are only executable if an Item is Selected
            PrintCommand = new RelayCommand(async _ => await PrintAsync(), _ => SelectedItem != null);
            VoidCommand = new RelayCommand(async _ => await VoidAsync(), _ => SelectedItem != null);
        }

        // --- Data Loading ---
        protected override async Task<List<AwrItemQueueDto>> FetchDataInternalAsync()
        {
            // Fetch items that are 'Issued' (Approved) and waiting for QC action
            return await Service.GetReceiptQueueAsync(Username);
        }

        // --- Selection Logic ---
        protected override void OnSelectionChanged()
        {
            // Clear input when selection changes
            VoidReason = string.Empty;

            // CRITICAL: Force WPF to re-evaluate CanExecute, 
            // which toggles the button IsEnabled, which triggers the Color Change.
            CommandManager.InvalidateRequerySuggested();
        }

        // --- Actions ---

        private async Task PrintAsync()
        {
            if (SelectedItem == null) return;

            if (MessageBox.Show($"Confirm printing for Request {SelectedItem.RequestNo}?\n\nThis will mark the document as 'Received'.",
                "Print Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.No)
                return;

            IsBusy = true;
            StatusMessage = "Processing Print Job...";

            try
            {
                // 1. Trigger Worker (Silent Print) + DB Update
                await Service.PrintAndReceiveItemAsync(SelectedItem.ItemId, Username);

                MessageBox.Show("Document sent to printer successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);

                // 2. Refresh Grid to remove processed item
                await LoadDataAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Print Failed: {ex.Message}\n\nEnsure Awr.Worker.exe is in the application folder.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsBusy = false;
                StatusMessage = "Ready";
            }
        }

        private async Task VoidAsync()
        {
            if (SelectedItem == null) return;

            // Validation
            if (string.IsNullOrWhiteSpace(VoidReason))
            {
                MessageBox.Show("Please enter a 'Void Reason' before cancelling.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (MessageBox.Show($"Are you sure you want to VOID Request {SelectedItem.RequestNo}?\n\nThis action cannot be undone.",
                "Confirm Void", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.No)
                return;

            IsBusy = true;
            StatusMessage = "Voiding Request...";

            try
            {
                // 1. Update DB (Mark as Returned/Voided)
                await Service.VoidItemAsync(SelectedItem.ItemId, Username, VoidReason);

                MessageBox.Show("Request has been Voided.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);

                // 2. Refresh Grid
                await LoadDataAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Void Failed: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsBusy = false;
                StatusMessage = "Ready";
            }
        }
    }
}