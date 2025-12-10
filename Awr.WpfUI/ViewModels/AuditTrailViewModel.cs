using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows; // For MessageBox
using System.Windows.Input;
using Awr.Core.DTOs;
using Awr.Core.Enums;
using Awr.WpfUI.MvvmCore;
using Awr.WpfUI.Services.Implementation; // ReportService
using Awr.WpfUI.ViewModels.Shared;

namespace Awr.WpfUI.ViewModels
{
    public class AuditTrailViewModel : WorkQueueViewModel
    {
        // --- Report Service ---
        private readonly ReportService _reportService = new ReportService();

        // --- Search Fields ---
        private string _searchArNo;
        public string SearchArNo { get => _searchArNo; set { if (SetProperty(ref _searchArNo, value)) _ = LoadDataAsync(); } }

        private bool _showAll = true;
        public bool ShowAll { get => _showAll; set { if (SetProperty(ref _showAll, value)) _ = LoadDataAsync(); } }

        // --- PAGING Logic ---
        public int PageSize { get; } = 50;
        private int _currentPage = 1;
        public int CurrentPage { get => _currentPage; set { SetProperty(ref _currentPage, value); CommandManager.InvalidateRequerySuggested(); } }
        private int _totalPages = 1;
        public int TotalPages { get => _totalPages; set => SetProperty(ref _totalPages, value); }
        private bool _isPagingEnabled = true;
        public bool IsPagingEnabled { get => _isPagingEnabled; set => SetProperty(ref _isPagingEnabled, value); }
        private int _totalRecords;
        public int TotalRecords { get => _totalRecords; set => SetProperty(ref _totalRecords, value); }

        // --- Status Logic ---
        public class StatusOption { public string Value { get; set; } public string Display { get; set; } }
        public ObservableCollection<StatusOption> StatusOptions { get; } = new ObservableCollection<StatusOption>();
        private StatusOption _selectedStatusOption;
        public StatusOption SelectedStatusOption { get => _selectedStatusOption; set { if (SetProperty(ref _selectedStatusOption, value)) _ = LoadDataAsync(); } }

        // --- Commands ---
        public ICommand NextPageCommand { get; }
        public ICommand PrevPageCommand { get; }
        public ICommand ExportExcelCommand { get; } // NEW
        public ICommand ExportPdfCommand { get; }   // NEW

        public AuditTrailViewModel(string username) : base(username)
        {
            // Init Statuses
            StatusOptions.Add(new StatusOption { Value = "All", Display = "--- All Statuses ---" });
            foreach (AwrItemStatus s in Enum.GetValues(typeof(AwrItemStatus)))
            {
                if (s == AwrItemStatus.Draft || s == AwrItemStatus.Complete) continue;
                string display = s.ToString();
                switch (s)
                {
                    case AwrItemStatus.PendingIssuance: display = "Pending Approval"; break;
                    case AwrItemStatus.Issued: display = "Approved"; break;
                    case AwrItemStatus.Received: display = "Completed"; break;
                    case AwrItemStatus.Voided: display = "Voided"; break;
                    case AwrItemStatus.RejectedByQa: display = "Rejected"; break;
                }
                StatusOptions.Add(new StatusOption { Value = s.ToString(), Display = display });
            }
            SelectedStatusOption = StatusOptions[0];

            // Init Commands
            NextPageCommand = new RelayCommand(async _ => { CurrentPage++; await LoadDataAsync(); }, _ => IsPagingEnabled && CurrentPage < TotalPages);
            PrevPageCommand = new RelayCommand(async _ => { CurrentPage--; await LoadDataAsync(); }, _ => IsPagingEnabled && CurrentPage > 1);
            
            // NEW: Export Commands
            ExportExcelCommand = new RelayCommand(_ => ExportData("Excel"));
            ExportPdfCommand = new RelayCommand(_ => ExportData("PDF"));
        }

        private void ExportData(string format)
        {
            // Use 'Items' (The currently visible/filtered list)
            var dataToExport = Items.ToList();
            if (!dataToExport.Any()) { MessageBox.Show("No data to export.", "Info"); return; }

            var dialog = new Microsoft.Win32.SaveFileDialog();
            dialog.FileName = $"AWR_AuditTrail_{DateTime.Now:yyyyMMdd_HHmm}";
            
            if (format == "Excel")
            {
                dialog.Filter = "Excel Workbook|*.xlsx";
                if (dialog.ShowDialog() == true)
                {
                    IsBusy = true;
                    // Run export on bg thread
                    Task.Run(() => _reportService.ExportToExcel(dataToExport, dialog.FileName, Username))
                        .ContinueWith(t => 
                        { 
                            IsBusy = false; 
                            if (t.IsFaulted) MessageBox.Show("Export Failed: " + t.Exception?.InnerException?.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                            else MessageBox.Show("Export Complete.", "Success", MessageBoxButton.OK, MessageBoxImage.Information); 
                        }, TaskScheduler.FromCurrentSynchronizationContext());
                }
            }
            else // PDF
            {
                dialog.Filter = "PDF Document|*.pdf";
                if (dialog.ShowDialog() == true)
                {
                    IsBusy = true;
                    Task.Run(() => _reportService.ExportToPdf(dataToExport, dialog.FileName, Username))
                         .ContinueWith(t => 
                         { 
                             IsBusy = false; 
                             if (t.IsFaulted) MessageBox.Show("Export Failed: " + t.Exception?.InnerException?.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                             else MessageBox.Show("Export Complete.", "Success", MessageBoxButton.OK, MessageBoxImage.Information); 
                         }, TaskScheduler.FromCurrentSynchronizationContext());
                }
            }
        }

        protected override async Task<List<AwrItemQueueDto>> FetchDataInternalAsync()
        {
            bool hasFilters = !string.IsNullOrEmpty(SearchText) || !string.IsNullOrEmpty(SearchArNo) 
                           || (SelectedStatusOption != null && SelectedStatusOption.Value != "All") || !ShowAll;

            if (hasFilters)
            {
                IsPagingEnabled = false; CurrentPage = 1; TotalPages = 1;
                var allData = await Service.GetAllAuditItemsAsync();
                TotalRecords = allData.Count;
                return allData;
            }
            else
            {
                IsPagingEnabled = true;
                return await Task.Run(() => 
                {
                    int total;
                    var list = Service.GetAuditItemsPaged(CurrentPage, PageSize, out total);
                    TotalRecords = total;
                    TotalPages = (int)Math.Ceiling((double)total / PageSize);
                    if (TotalPages < 1) TotalPages = 1;
                    return list;
                });
            }
        }

        protected override void FilterData()
        {
            if (AllItems == null) return;

            if (IsPagingEnabled)
            {
                Items = new ObservableCollection<AwrItemQueueDto>(AllItems);
            }
            else
            {
                var query = AllItems.AsEnumerable();

                if (!string.IsNullOrWhiteSpace(SearchText))
                {
                    string lower = SearchText.ToLower();
                    query = query.Where(i => 
                        (i.RequestNo?.ToLower().Contains(lower) ?? false) || 
                        (i.MaterialProduct?.ToLower().Contains(lower) ?? false) || 
                        (i.BatchNo?.ToLower().Contains(lower) ?? false) ||
                        (i.AwrNo?.ToLower().Contains(lower) ?? false)
                    );
                }

                if (!string.IsNullOrWhiteSpace(SearchArNo))
                {
                    var tokens = SearchArNo.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim().ToLower());
                    query = query.Where(i => 
                    {
                        string dbAr = i.ArNo?.ToLower() ?? "";
                        return tokens.Any(token => dbAr.Contains(token));
                    });
                }

                if (SelectedStatusOption != null && SelectedStatusOption.Value != "All")
                {
                    query = query.Where(i => i.Status.ToString() == SelectedStatusOption.Value);
                }

                if (!ShowAll)
                {
                    query = query.Where(i => string.Equals(i.RequestedBy, Username, StringComparison.OrdinalIgnoreCase));
                }

                var filteredList = query.ToList();
                Items = new ObservableCollection<AwrItemQueueDto>(filteredList);
                TotalRecords = filteredList.Count;
            }
        }
    }
}