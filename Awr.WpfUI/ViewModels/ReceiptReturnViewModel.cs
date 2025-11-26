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

        public ICommand PrintCommand { get; }
        public ICommand VoidCommand { get; }

        public ReceiptReturnViewModel(string username) : base(username)
        {
            PrintCommand = new RelayCommand(async _ => await PrintAsync(), _ => SelectedItem != null);
            VoidCommand = new RelayCommand(async _ => await VoidAsync(), _ => SelectedItem != null);
        }

        protected override async Task<List<AwrItemQueueDto>> FetchDataInternalAsync()
        {
            // QC sees items approved by QA waiting for print
            return await Service.GetReceiptQueueAsync(Username);
        }

        private async Task PrintAsync()
        {
            if (SelectedItem == null) return;
            if (MessageBox.Show($"Print Document for {SelectedItem.RequestNo}?", "Confirm", MessageBoxButton.YesNo) == MessageBoxResult.No) return;

            IsBusy = true;
            try
            {
                await Service.PrintAndReceiveItemAsync(SelectedItem.ItemId, Username);
                MessageBox.Show("Document sent to printer.", "Success");
                await LoadDataAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Print Failed: {ex.Message}\n\nCheck Worker/Printer.", "Error");
            }
            finally { IsBusy = false; }
        }

        private async Task VoidAsync()
        {
            if (SelectedItem == null) return;
            if (string.IsNullOrWhiteSpace(VoidReason))
            {
                MessageBox.Show("Void Reason is mandatory.", "Validation");
                return;
            }

            if (MessageBox.Show("Void/Cancel this request? This cannot be undone.", "Warning", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.No) return;

            IsBusy = true;
            try
            {
                await Service.VoidItemAsync(SelectedItem.ItemId, Username, VoidReason);
                MessageBox.Show("Request Voided.", "Success");
                VoidReason = string.Empty;
                await LoadDataAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Void Failed: {ex.Message}", "Error");
            }
            finally { IsBusy = false; }
        }
    }
}