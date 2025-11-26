using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using Awr.Core.DTOs;
using Awr.Core.Enums;
using Awr.WpfUI.MvvmCore;
using Awr.WpfUI.Services.Implementation;
using Awr.WpfUI.Services.Interfaces;

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
        public bool IsBusy { get => _isBusy; set { SetProperty(ref _isBusy, value); CommandManager.InvalidateRequerySuggested(); } }

        private string _statusMessage;
        public string StatusMessage { get => _statusMessage; set => SetProperty(ref _statusMessage, value); }

        public ICommand SubmitCommand { get; }
        public ICommand ReloadSequenceCommand { get; }

        public RequestSubmissionViewModel() { /* Design-time only */ }

        public RequestSubmissionViewModel(string username)
        {
            _username = username;
            _service = new WorkflowService(); // Manual DI

            SelectedType = AwrTypes.FirstOrDefault();

            SubmitCommand = new RelayCommand(async _ => await SubmitAsync(), _ => !IsBusy);
            ReloadSequenceCommand = new RelayCommand(_ => LoadNextSequence());

            LoadNextSequence();
        }

        private void LoadNextSequence()
        {
            try
            {
                string seq = _service.GetNextRequestNumberSequenceValue();
                string datePart = DateTime.Now.ToString("yyyyMMdd");
                RequestNo = $"AWR-{datePart}-{seq}";
                StatusMessage = "Ready.";
            }
            catch (Exception ex)
            {
                RequestNo = "ERROR";
                StatusMessage = "Failed to load sequence: " + ex.Message;
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

                await _service.SubmitNewRequestAsync(dto, _username, RequestNo);

                MessageBox.Show($"Request {RequestNo} submitted successfully.", "Success");
                ClearForm();
                LoadNextSequence();
            }
            catch (Exception ex)
            {
                StatusMessage = "Error";
                MessageBox.Show($"Submission Failed: {ex.Message}", "Error");
            }
            finally
            {
                IsBusy = false;
            }
        }

        private bool ValidateForm()
        {
            if (string.IsNullOrWhiteSpace(MaterialProduct)) { MessageBox.Show("Material/Product is required."); return false; }
            if (string.IsNullOrWhiteSpace(BatchNo)) { MessageBox.Show("Batch No is required."); return false; }
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