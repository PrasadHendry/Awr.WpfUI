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
    public class IssuanceQueueViewModel : WorkQueueViewModel
    {
        private string _rejectComment;
        public string RejectComment { get => _rejectComment; set => SetProperty(ref _rejectComment, value); }

        public ICommand ApproveCommand { get; }
        public ICommand RejectCommand { get; }

        public IssuanceQueueViewModel(string username) : base(username)
        {
            ApproveCommand = new RelayCommand(async _ => await ApproveAsync(), _ => SelectedItem != null);
            RejectCommand = new RelayCommand(async _ => await RejectAsync(), _ => SelectedItem != null);
        }

        protected override async Task<List<AwrItemQueueDto>> FetchDataInternalAsync()
        {
            return await Service.GetIssuanceQueueAsync();
        }

        private async Task ApproveAsync()
        {
            if (SelectedItem == null) return;
            if (MessageBox.Show($"Approve Request {SelectedItem.RequestNo}?", "Confirm", MessageBoxButton.YesNo) == MessageBoxResult.No) return;

            IsBusy = true;
            try
            {
                await Service.IssueItemAsync(SelectedItem.ItemId, Username);
                MessageBox.Show("Approved & Generated successfully.", "Success");
                await LoadDataAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Approval Failed: {ex.Message}\n\nEnsure Awr.Worker.exe is in the bin folder.", "Error");
            }
            finally { IsBusy = false; }
        }

        private async Task RejectAsync()
        {
            if (SelectedItem == null) return;
            if (string.IsNullOrWhiteSpace(RejectComment))
            {
                MessageBox.Show("Rejection comment is mandatory.", "Validation");
                return;
            }

            if (MessageBox.Show("Reject this request?", "Confirm", MessageBoxButton.YesNo) == MessageBoxResult.No) return;

            IsBusy = true;
            try
            {
                await Service.RejectItemAsync(SelectedItem.ItemId, Username, RejectComment);
                MessageBox.Show("Request Rejected.", "Success");
                RejectComment = string.Empty; // Reset
                await LoadDataAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Rejection Failed: {ex.Message}", "Error");
            }
            finally { IsBusy = false; }
        }
    }
}