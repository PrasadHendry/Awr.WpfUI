using Awr.Core.DTOs;
using Awr.WpfUI.MvvmCore;
using Awr.WpfUI.ViewModels.Shared;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace Awr.WpfUI.ViewModels
{
    public class IssuanceQueueViewModel : WorkQueueViewModel
    {
        // --- State ---
        private string _rejectComment;
        public string RejectComment
        {
            get => _rejectComment;
            set => SetProperty(ref _rejectComment, value);
        }

        // --- Commands ---
        public ICommand ApproveCommand { get; }
        public ICommand RejectCommand { get; }

        // --- Constructor ---
        public IssuanceQueueViewModel(string username) : base(username)
        {
            // Initialize Commands
            // CanExecute logic ensures buttons are only active if an item is selected
            ApproveCommand = new RelayCommand(async _ => await ApproveAsync(), _ => SelectedItem != null);
            RejectCommand = new RelayCommand(async _ => await RejectAsync(), _ => SelectedItem != null);
        }

        // --- Data Loading ---
        protected override async Task<List<AwrItemQueueDto>> FetchDataInternalAsync()
        {
            // Fetch only items waiting for QA approval
            return await Service.GetIssuanceQueueAsync();
        }

        // --- Selection Logic ---
        protected override void OnSelectionChanged()
        {
            // When user picks a new row, clear the rejection box
            RejectComment = string.Empty;

            // Force UI to re-evaluate button enablement
            CommandManager.InvalidateRequerySuggested();
        }

        // --- Actions ---

        private async Task ApproveAsync()
        {
            if (SelectedItem == null) return;

            // 1. CHECK DUPLICATE AR
            IsBusy = true;
            try
            {
                // Run on bg thread
                List<string> duplicates = await Task.Run(() => Service.CheckIfArNumberExists(SelectedItem.ArNo.Trim()));

                if (duplicates.Any())
                {
                    string msg = "WARNING: Duplicate AR Number(s) found:\n\n";
                    msg += string.Join("\n", duplicates.Take(10));
                    msg += "\n\nDo you want to Approve anyway?";

                    if (MessageBox.Show(msg, "Duplicate", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.No)
                    {
                        IsBusy = false;
                        StatusMessage = "Cancelled";
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Duplicate check failed: " + ex.Message);
                IsBusy = false;
                return;
            }

            // 2. STANDARD CONFIRMATION
            string confirmMsg = $"Approve Request {SelectedItem.RequestNo}?";
            if (SelectedItem.QtyRequired > 1)
            {
                confirmMsg += $"\n\nQuantity to Issue: {SelectedItem.QtyRequired:0}";
            }

            if (MessageBox.Show(confirmMsg, "Confirm Approval", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.No)
            {
                IsBusy = false;
                return;
            }

            // 3. EXECUTE APPROVAL
            StatusMessage = "Generating Document...";
            try
            {
                await Service.IssueItemAsync(SelectedItem.ItemId, SelectedItem.QtyRequired, Username);
                MessageBox.Show($"Request approved.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                await LoadDataAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Approval Failed: {ex.Message}\n\nEnsure Worker is present.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsBusy = false;
                StatusMessage = "Ready";
            }
        }

        private async Task RejectAsync()
        {
            if (SelectedItem == null) return;

            // Validation: Comment is mandatory for rejection
            if (string.IsNullOrWhiteSpace(RejectComment))
            {
                MessageBox.Show("A 'Rejection Reason' is required to reject a request.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (MessageBox.Show($"Are you sure you want to REJECT Request {SelectedItem.RequestNo}?",
                "Confirm Rejection", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.No)
                return;

            IsBusy = true;
            StatusMessage = "Rejecting...";

            try
            {
                // 1. Call Service (DB Update only)
                await Service.RejectItemAsync(SelectedItem.ItemId, Username, RejectComment);

                MessageBox.Show("Request Rejected.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);

                // 2. Refresh Grid
                await LoadDataAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Rejection Failed: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsBusy = false;
                StatusMessage = "Ready";
            }
        }
    }
}