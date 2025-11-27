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
    public class NewRequestViewModel : BaseViewModel
    {
        private readonly IWorkflowService _service;
        private readonly string _username;

        // --- Data ---
        public string RequestNo { get => _requestNo; set => SetProperty(ref _requestNo, value); }
        private string _requestNo;

        public AwrType SelectedType { get => _selectedType; set => SetProperty(ref _selectedType, value); }
        private AwrType _selectedType;
        public IEnumerable<AwrType> AwrTypes => Enum.GetValues(typeof(AwrType)).Cast<AwrType>().Where(t => t != AwrType.Others);

        public string MaterialProduct { get => _materialProduct; set => SetProperty(ref _materialProduct, value); }
        private string _materialProduct;

        public string BatchNo { get => _batchNo; set => SetProperty(ref _batchNo, value); }
        private string _batchNo;

        public string ArNo { get => _arNo; set => SetProperty(ref _arNo, value); }
        private string _arNo;

        public string AwrNo { get => _awrNo; set => SetProperty(ref _awrNo, value); }
        private string _awrNo;

        public string Comments { get => _comments; set => SetProperty(ref _comments, value); }
        private string _comments;

        // --- Error States ---
        private bool _isMaterialError;
        public bool IsMaterialError { get => _isMaterialError; set => SetProperty(ref _isMaterialError, value); }

        private bool _isBatchError;
        public bool IsBatchError { get => _isBatchError; set => SetProperty(ref _isBatchError, value); }

        private bool _isAwrError;
        public bool IsAwrError { get => _isAwrError; set => SetProperty(ref _isAwrError, value); }

        // --- NEW ERROR STATES ---
        private bool _isArError;
        public bool IsArError { get => _isArError; set => SetProperty(ref _isArError, value); }

        private bool _isCommentError;
        public bool IsCommentError { get => _isCommentError; set => SetProperty(ref _isCommentError, value); }

        // --- State ---
        public bool IsBusy { get => _isBusy; set { SetProperty(ref _isBusy, value); CommandManager.InvalidateRequerySuggested(); } }
        private bool _isBusy;

        public string StatusMessage { get => _statusMessage; set => SetProperty(ref _statusMessage, value); }
        private string _statusMessage;

        public ICommand SubmitCommand { get; }

        public NewRequestViewModel() { }

        public NewRequestViewModel(string username)
        {
            _username = username;
            _service = new WorkflowService();
            SelectedType = AwrTypes.FirstOrDefault();
            SubmitCommand = new RelayCommand(async _ => await SubmitAsync(), _ => !IsBusy);
            LoadNextSequence();
        }

        private void LoadNextSequence()
        {
            try
            {
                string seq = _service.GetNextRequestNumberSequenceValue();
                RequestNo = $"AWR-{DateTime.Now:yyyyMMdd}-{seq}";
                StatusMessage = "Ready for submission.";
            }
            catch (Exception ex) { RequestNo = "ERROR"; StatusMessage = "Error: " + ex.Message; }
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
                    Items = new List<AwrItemSubmissionDto> { new AwrItemSubmissionDto { MaterialProduct = MaterialProduct, BatchNo = BatchNo, ArNo = ArNo, AwrNo = AwrNo, QtyRequired = 1 } }
                };

                await _service.SubmitNewRequestAsync(dto, _username, RequestNo);
                MessageBox.Show($"Request {RequestNo} submitted!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                ClearForm();
                LoadNextSequence();
            }
            catch (Exception ex)
            {
                StatusMessage = "Failed.";
                MessageBox.Show(ex.Message, "Error");
            }
            finally { IsBusy = false; }
        }

        private bool ValidateForm()
        {
            IsMaterialError = false; IsBatchError = false; IsAwrError = false;
            IsArError = false; IsCommentError = false; // Reset new flags

            bool isValid = true;

            if (string.IsNullOrWhiteSpace(MaterialProduct)) { IsMaterialError = true; isValid = false; }
            if (string.IsNullOrWhiteSpace(BatchNo)) { IsBatchError = true; isValid = false; }
            if (string.IsNullOrWhiteSpace(AwrNo)) { IsAwrError = true; isValid = false; }

            // NEW VALIDATIONS
            if (string.IsNullOrWhiteSpace(ArNo)) { IsArError = true; isValid = false; }
            if (string.IsNullOrWhiteSpace(Comments)) { IsCommentError = true; isValid = false; }

            if (!isValid) MessageBox.Show("Please fill in the required highlighted fields.", "Validation");
            return isValid;
        }

        private void ClearForm()
        {
            MaterialProduct = ""; BatchNo = ""; ArNo = ""; AwrNo = ""; Comments = "";
            IsMaterialError = false; IsBatchError = false; IsAwrError = false;
            IsArError = false; IsCommentError = false;
            SelectedType = AwrTypes.FirstOrDefault();
        }
    }
}