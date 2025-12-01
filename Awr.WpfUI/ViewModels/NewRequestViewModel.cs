using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Configuration;
using System.IO;
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
        private readonly string _csvPath;
        private List<string> _masterAwrList = new List<string>();

        // --- Data ---
        public string RequestNo { get => _requestNo; set => SetProperty(ref _requestNo, value); }
        private string _requestNo;

        // --- TYPE LOGIC ---
        public AwrType SelectedType 
        { 
            get => _selectedType; 
            set 
            { 
                if (SetProperty(ref _selectedType, value))
                {
                    if (_selectedType == AwrType.RM) { IsQtyEnabled = true; }
                    else { IsQtyEnabled = false; QtyRequired = 1; }
                }
            } 
        }
        private AwrType _selectedType;
        public IEnumerable<AwrType> AwrTypes => Enum.GetValues(typeof(AwrType)).Cast<AwrType>().Where(t => t != AwrType.Others);

        // --- SEARCHABLE AWR ---
        public ObservableCollection<string> FilteredAwrNumbers { get; private set; } = new ObservableCollection<string>();
        public string AwrNo 
        { 
            get => _awrNo; 
            set { if (SetProperty(ref _awrNo, value)) FilterAwrList(); } 
        }
        private string _awrNo;

        // --- QUANTITY ---
        public decimal QtyRequired { get => _qtyRequired; set => SetProperty(ref _qtyRequired, value); }
        private decimal _qtyRequired = 1;
        public bool IsQtyEnabled { get => _isQtyEnabled; set => SetProperty(ref _isQtyEnabled, value); }
        private bool _isQtyEnabled;

        // --- FORM FIELDS ---
        public string MaterialProduct { get => _materialProduct; set => SetProperty(ref _materialProduct, value); }
        private string _materialProduct;
        public string BatchNo { get => _batchNo; set => SetProperty(ref _batchNo, value); }
        private string _batchNo;
        public string ArNo { get => _arNo; set => SetProperty(ref _arNo, value); }
        private string _arNo;
        public string Comments { get => _comments; set => SetProperty(ref _comments, value); }
        private string _comments;

        // --- ERROR STATES ---
        private bool _isMaterialError; public bool IsMaterialError { get => _isMaterialError; set => SetProperty(ref _isMaterialError, value); }
        private bool _isBatchError; public bool IsBatchError { get => _isBatchError; set => SetProperty(ref _isBatchError, value); }
        private bool _isAwrError; public bool IsAwrError { get => _isAwrError; set => SetProperty(ref _isAwrError, value); }
        private bool _isArError; public bool IsArError { get => _isArError; set => SetProperty(ref _isArError, value); }
        private bool _isCommentError; public bool IsCommentError { get => _isCommentError; set => SetProperty(ref _isCommentError, value); }

        // --- STATE ---
        public bool IsBusy { get => _isBusy; set { SetProperty(ref _isBusy, value); CommandManager.InvalidateRequerySuggested(); } }
        private bool _isBusy;
        public string StatusMessage { get => _statusMessage; set => SetProperty(ref _statusMessage, value); }
        private string _statusMessage;

        // --- COMMANDS ---
        public ICommand SubmitCommand { get; }
        public ICommand IncrementQtyCommand { get; }
        public ICommand DecrementQtyCommand { get; }

        public NewRequestViewModel() { }

        public NewRequestViewModel(string username)
        {
            _username = username;
            _service = new WorkflowService();
            _csvPath = ConfigurationManager.AppSettings["AwrMasterCsvPath"];

            SelectedType = AwrTypes.FirstOrDefault();
            
            SubmitCommand = new RelayCommand(async _ => await SubmitAsync(), _ => !IsBusy);
            IncrementQtyCommand = new RelayCommand(_ => QtyRequired++, _ => IsQtyEnabled);
            DecrementQtyCommand = new RelayCommand(_ => { if (QtyRequired > 1) QtyRequired--; }, _ => IsQtyEnabled && QtyRequired > 1);

            LoadNextSequence();
            LoadAwrFromCsv();
        }

        private void LoadAwrFromCsv()
        {
            try
            {
                if (!string.IsNullOrEmpty(_csvPath) && File.Exists(_csvPath))
                {
                    var lines = File.ReadAllLines(_csvPath);
                    // Skip Header, take first column, TRIM QUOTES and SPACES
                    _masterAwrList = lines.Skip(1)
                                          .Select(x => x.Split(',')[0].Trim().Trim('"')) // <--- Added Trim('"')
                                          .Where(x => !string.IsNullOrWhiteSpace(x))
                                          .OrderBy(x => x)
                                          .ToList();
                }
                else
                {
                    // Fallback for testing
                    _masterAwrList = new List<string> { "AWR-TEST-001", "AWR-TEST-002" };
                }

                FilterAwrList();
            }
            catch (Exception ex) { MessageBox.Show("Error loading AWR CSV: " + ex.Message); }
        }

        private void FilterAwrList()
        {
            if (_masterAwrList == null) return;

            string search = AwrNo?.ToLower() ?? "";

            // FIX: If the text is an EXACT match to an item, assume the user selected it 
            // and do not filter down (or reset to full list to be safe).
            if (_masterAwrList.Any(x => x.Equals(AwrNo, StringComparison.OrdinalIgnoreCase)))
            {
                // Optional: You could reset to full list here, or just do nothing.
                return;
            }

            var matches = _masterAwrList.Where(x => x.ToLower().Contains(search)).ToList();

            FilteredAwrNumbers.Clear();
            foreach (var item in matches) FilteredAwrNumbers.Add(item);
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
                    Items = new List<AwrItemSubmissionDto> 
                    { 
                        new AwrItemSubmissionDto { MaterialProduct = MaterialProduct, BatchNo = BatchNo, ArNo = ArNo, AwrNo = AwrNo, QtyRequired = QtyRequired } 
                    }
                };
                await _service.SubmitNewRequestAsync(dto, _username, RequestNo);
                MessageBox.Show($"Request {RequestNo} submitted!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                ClearForm();
                LoadNextSequence();
            }
            catch (Exception ex) { StatusMessage = "Failed."; MessageBox.Show(ex.Message, "Error"); }
            finally { IsBusy = false; }
        }

        private bool ValidateForm()
        {
            var missingFields = new List<string>();

            if (string.IsNullOrWhiteSpace(MaterialProduct)) missingFields.Add("- Material");
            if (string.IsNullOrWhiteSpace(BatchNo)) missingFields.Add("- Batch No");
            if (string.IsNullOrWhiteSpace(ArNo)) missingFields.Add("- AR No");
            if (string.IsNullOrWhiteSpace(Comments)) missingFields.Add("- Comments");

            // Validate AWR No
            if (string.IsNullOrWhiteSpace(AwrNo))
            {
                missingFields.Add("- AWR No");
            }
            else if (!_masterAwrList.Contains(AwrNo))
            {
                MessageBox.Show($"The selected AWR No '{AwrNo}' is invalid.\nPlease select a value from the list.", "Invalid Input");
                return false;
            }

            if (QtyRequired <= 0) missingFields.Add("- Quantity (must be > 0)");

            // If there are missing fields, show specific message
            if (missingFields.Any())
            {
                string message = "The following required fields are missing:\n\n" + string.Join("\n", missingFields);
                MessageBox.Show(message, "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            return true;
        }

        private void ClearForm()
        {
            MaterialProduct = ""; BatchNo = ""; ArNo = ""; AwrNo = ""; Comments = "";
            IsMaterialError = false; IsBatchError = false; IsAwrError = false; IsArError = false; IsCommentError = false;
            SelectedType = AwrTypes.FirstOrDefault();
            FilterAwrList();
        }
    }
}