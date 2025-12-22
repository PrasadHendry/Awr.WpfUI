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

        private class AwrCsvRecord
        {
            public string FileName { get; set; }
            public string ParentFolder { get; set; }
        }
        private List<AwrCsvRecord> _masterRecords = new List<AwrCsvRecord>();

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

                    // Reset selection & Refilter
                    AwrNo = "";
                    FilterAwrList();
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
            set
            {
                // Only filter if value changed
                if (SetProperty(ref _awrNo, value))
                {
                    FilterAwrList();
                }
            }
        }
        private string _awrNo = "";

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

        // --- ERRORS ---
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

            // Init Default Type (Triggering logic will happen after load)
            _selectedType = AwrTypes.FirstOrDefault();

            SubmitCommand = new RelayCommand(async _ => await SubmitAsync(), _ => !IsBusy);
            IncrementQtyCommand = new RelayCommand(_ => QtyRequired++, _ => IsQtyEnabled);
            DecrementQtyCommand = new RelayCommand(_ => { if (QtyRequired > 1) QtyRequired--; }, _ => IsQtyEnabled && QtyRequired > 1);

            ResetFormState();
            LoadAwrFromCsv();
        }

        private void ResetFormState()
        {
            RequestNo = $"AWR-{DateTime.Now:yyyyMMdd}-####";
            StatusMessage = "Ready for submission.";
        }

        private void LoadAwrFromCsv()
        {
            try
            {
                if (!string.IsNullOrEmpty(_csvPath) && File.Exists(_csvPath))
                {
                    var lines = File.ReadAllLines(_csvPath);
                    _masterRecords.Clear();

                    foreach (var line in lines.Skip(1))
                    {
                        if (string.IsNullOrWhiteSpace(line)) continue;
                        var parts = line.Split(',');

                        if (parts.Length >= 2)
                        {
                            string fileName = parts[0].Trim().Trim('"');
                            string parentFolder = "";

                            // FIX: Scan backwards to find the last non-empty column.
                            // This handles cases where CSV has trailing empty commas (e.g. "Name, Path, Folder, ,")
                            for (int i = parts.Length - 1; i > 0; i--)
                            {
                                if (!string.IsNullOrWhiteSpace(parts[i]))
                                {
                                    parentFolder = parts[i].Trim().Trim('"');
                                    break;
                                }
                            }

                            _masterRecords.Add(new AwrCsvRecord
                            {
                                FileName = fileName,
                                ParentFolder = parentFolder
                            });
                        }
                    }
                }
                else { _masterRecords = new List<AwrCsvRecord>(); }

                // Trigger Initial Filter
                FilterAwrList();
            }
            catch (Exception ex) { MessageBox.Show("Error loading CSV: " + ex.Message); }
        }

        private void FilterAwrList()
        {
            if (_masterRecords == null) return;

            string search = AwrNo?.ToLower() ?? "";

            string folderKeyword = "";
            switch (SelectedType)
            {
                case AwrType.FPS: folderKeyword = "FPS"; break;
                case AwrType.IMS: folderKeyword = "FPS"; break;
                case AwrType.MICRO: folderKeyword = "Micro"; break;
                case AwrType.PM: folderKeyword = "PM"; break;
                case AwrType.RM: folderKeyword = "RM"; break;
                case AwrType.STABILITY: folderKeyword = "Stability"; break;
                case AwrType.WATER: folderKeyword = "Water"; break;
            }

            // Optimization: If current text matches a valid file exactly, don't filter
            if (_masterRecords.Any(r => r.FileName.Equals(AwrNo, StringComparison.OrdinalIgnoreCase))) return;

            var query = _masterRecords.AsEnumerable();

            // 1. Folder Keyword Check
            if (!string.IsNullOrEmpty(folderKeyword))
            {
                query = query.Where(r => r.ParentFolder != null &&
                                         r.ParentFolder.IndexOf(folderKeyword, StringComparison.OrdinalIgnoreCase) >= 0);
            }

            // 2. Search Text
            if (!string.IsNullOrEmpty(search))
            {
                query = query.Where(r => r.FileName.ToLower().Contains(search));
            }

            var results = query.Select(r => r.FileName).OrderBy(x => x).ToList();

            FilteredAwrNumbers.Clear();
            if (results.Count == 0)
            {
                FilteredAwrNumbers.Add($"(No files found for {SelectedType})");
            }
            else
            {
                foreach (var item in results) FilteredAwrNumbers.Add(item);
            }
        }

        private async Task SubmitAsync()
        {
            if (!ValidateForm()) return;
            IsBusy = true;
            try
            {
                List<string> duplicates = await Task.Run(() => _service.CheckIfArNumberExists(ArNo.Trim(), null));

                if (duplicates.Any())
                {
                    string msg = "WARNING: The AR Number(s) are active in previous requests:\n\n";
                    msg += string.Join("\n", duplicates.Take(10));
                    if (duplicates.Count > 10) msg += "\n...";
                    msg += "\n\nDo you want to continue?";
                    if (MessageBox.Show(msg, "Duplicate Detected", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.No)
                    {
                        IsBusy = false; return;
                    }
                }

                StatusMessage = "Submitting...";

                var dto = new AwrRequestSubmissionDto
                {
                    Type = SelectedType,
                    RequestComment = Comments,
                    Items = new List<AwrItemSubmissionDto>
                    {
                        new AwrItemSubmissionDto { MaterialProduct = MaterialProduct, BatchNo = BatchNo, ArNo = ArNo, AwrNo = AwrNo, QtyRequired = QtyRequired }
                    }
                };

                string finalId = await _service.SubmitNewRequestAsync(dto, _username, "AUTO");

                MessageBox.Show($"Request Created Successfully!\n\nID: {finalId}", "Success", MessageBoxButton.OK, MessageBoxImage.Information);

                RequestNo = finalId;
                ClearForm();
                ResetFormState();
            }
            catch (Exception ex)
            {
                StatusMessage = "Failed.";
                MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsBusy = false;
                if (StatusMessage == "Submitting...") StatusMessage = "Ready for submission.";
            }
        }

        private bool ValidateForm()
        {
            var missingFields = new List<string>();

            if (string.IsNullOrWhiteSpace(MaterialProduct)) { IsMaterialError = true; missingFields.Add("- Material/Product"); }
            if (string.IsNullOrWhiteSpace(BatchNo)) { IsBatchError = true; missingFields.Add("- Batch No"); }
            if (string.IsNullOrWhiteSpace(ArNo)) { IsArError = true; missingFields.Add("- AR No"); }
            if (string.IsNullOrWhiteSpace(Comments)) { IsCommentError = true; missingFields.Add("- Comments"); }

            if (string.IsNullOrWhiteSpace(AwrNo))
            {
                IsAwrError = true; missingFields.Add("- AWR No");
            }
            else if (!_masterRecords.Any(r => r.FileName == AwrNo))
            {
                MessageBox.Show($"The selected AWR No '{AwrNo}' is invalid.", "Invalid Input", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            if (QtyRequired <= 0) missingFields.Add("- Quantity (must be > 0)");

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