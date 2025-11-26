using Awr.Core.DTOs;
using Awr.Core.Enums;
using Awr.WpfUI.MvvmCore;
using Awr.WpfUI.Services.Implementation;
using Awr.WpfUI.Services.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace Awr.WpfUI.ViewModels
{
    public class RequestSubmissionViewModel : BaseViewModel
    {
        private readonly IWorkflowService _service;
        private readonly string _username;

        // --- Form Properties ---
        private string _requestNo;
        public string RequestNo { get => _requestNo; set => SetProperty(ref _requestNo, value); }

        private AwrType _selectedType;
        public AwrType SelectedType { get => _selectedType; set => SetProperty(ref _selectedType, value); }

        // Source for Dropdown
        public IEnumerable<AwrType> AwrTypes => Enum.GetValues(typeof(AwrType)).Cast<AwrType>().Where(t => t != AwrType.Others);

        private string _materialProduct;
        public string MaterialProduct { get => _materialProduct; set => SetProperty(ref _materialProduct, value); }

        private string _batchNo;
        public string BatchNo { get => _batchNo; set => SetProperty(ref _batchNo, value); }

        private string _arNo;
        public string ArNo { get => _arNo; set => SetProperty(ref _arNo, value); }

        private string _awrNo;
        public string AwrNo { get => _awrNo; set => SetProperty(ref _awrNo, value); }

        private string _comments;
        public string Comments { get => _comments; set => SetProperty(ref _comments, value); }

        // --- State ---
        private bool _isBusy;
        public bool IsBusy
        {
            get => _isBusy;
            set
            {
                SetProperty(ref _isBusy, value);
                CommandManager.InvalidateRequerySuggested();
            }
        }

        private string _statusMessage;
        public string StatusMessage { get => _statusMessage; set => SetProperty(ref _statusMessage, value); }

        public ICommand SubmitCommand { get; }

        // Design-time Constructor
        public RequestSubmissionViewModel() { }

        // Runtime Constructor
        public RequestSubmissionViewModel(string username)
        {
            _username = username;
            _service = new WorkflowService(); // Ideally use DI

            SelectedType = AwrTypes.FirstOrDefault();
            StatusMessage = "Initializing...";

            SubmitCommand = new RelayCommand(async _ => await SubmitAsync(), _ => !IsBusy);

            // Generate ID on load
            LoadNextSequence();
        }

        private void LoadNextSequence()
        {
            try
            {
                // Synchronous call is fine for simple scalar
                string seq = _service.GetNextRequestNumberSequenceValue();
                string datePart = DateTime.Now.ToString("yyyyMMdd");
                RequestNo = $"AWR-{datePart}-{seq}";
                StatusMessage = "Ready for submission.";
            }
            catch (Exception ex)
            {
                RequestNo = "ERROR";
                StatusMessage = "Error: " + ex.Message;
            }
        }

        private async Task SubmitAsync()
        {
            if (!ValidateForm()) return;

            IsBusy = true;
            StatusMessage = "Submitting...";

            try
            {
                var dto = new AwrRequestSubmissionDto
                {
                    Type = SelectedType,
                    RequestComment = Comments,
                    Items = new List<AwrItemSubmissionDto>
                    {
                        new AwrItemSubmissionDto
                        {
                            MaterialProduct = MaterialProduct,
                            BatchNo = BatchNo,
                            ArNo = ArNo,
                            AwrNo = AwrNo,
                            QtyRequired = 1
                        }
                    }
                };

                // Call Service
                await _service.SubmitNewRequestAsync(dto, _username, RequestNo);

                MessageBox.Show($"Request {RequestNo} submitted successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);

                // Reset Form
                ClearForm();
                LoadNextSequence(); // Get NEW ID for next request
            }
            catch (Exception ex)
            {
                StatusMessage = "Submission Failed.";
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsBusy = false;
            }
        }

        private bool ValidateForm()
        {
            if (string.IsNullOrWhiteSpace(MaterialProduct)) { MessageBox.Show("Material is required."); return false; }
            if (string.IsNullOrWhiteSpace(BatchNo)) { MessageBox.Show("Batch No is required."); return false; }
            // AWR No is required in WinForms logic? If so, keep this check.
            if (string.IsNullOrWhiteSpace(AwrNo)) { MessageBox.Show("AWR No is required."); return false; }
            return true;
        }

        private void ClearForm()
        {
            MaterialProduct = string.Empty;
            BatchNo = string.Empty;
            ArNo = string.Empty;
            AwrNo = string.Empty;
            Comments = string.Empty;
            SelectedType = AwrTypes.FirstOrDefault();
        }
    }
}