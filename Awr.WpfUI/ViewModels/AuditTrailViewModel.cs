using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using Awr.Core.DTOs;
using Awr.Core.Enums;
using Awr.WpfUI.MvvmCore;
using Awr.WpfUI.Services.Implementation;
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

        // Holds all filtered items locally for Exporting
        private List<AwrItemQueueDto> _filteredItemsList = new List<AwrItemQueueDto>();

        // --- Status Logic ---
        public class StatusOption { public string Value { get; set; } public string Display { get; set; } }
        public ObservableCollection<StatusOption> StatusOptions { get; } = new ObservableCollection<StatusOption>();
        private StatusOption _selectedStatusOption;
        public StatusOption SelectedStatusOption { get => _selectedStatusOption; set { if (SetProperty(ref _selectedStatusOption, value)) _ = LoadDataAsync(); } }

        // --- Commands ---
        public ICommand NextPageCommand { get; }
        public ICommand PrevPageCommand { get; }
        public ICommand ExportExcelCommand { get; }
        public ICommand ExportPdfCommand { get; }

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

            // Init Commands - Paging operations now happen strictly locally
            NextPageCommand = new RelayCommand(_ => { CurrentPage++; FilterData(); }, _ => IsPagingEnabled && CurrentPage < TotalPages);
            PrevPageCommand = new RelayCommand(_ => { CurrentPage--; FilterData(); }, _ => IsPagingEnabled && CurrentPage > 1);

            // Init Export Commands
            ExportExcelCommand = new RelayCommand(_ => ExportData("Excel"));
            ExportPdfCommand = new RelayCommand(_ => ExportData("PDF"));
        }

        private void ExportData(string format)
        {
            // Export the full filtered list, not just the current page
            var dataToExport = (_filteredItemsList != null && _filteredItemsList.Any()) ? _filteredItemsList : Items.ToList();

            if (dataToExport == null || !dataToExport.Any())
            {
                MessageBox.Show("No data to export.", "Info");
                return;
            }

            var dialog = new Microsoft.Win32.SaveFileDialog();
            dialog.FileName = $"AWR_AuditTrail_{DateTime.Now:yyyyMMdd_HHmm}";

            if (format == "Excel")
            {
                dialog.Filter = "Excel Workbook|*.xlsx";
                if (dialog.ShowDialog() == true)
                {
                    IsBusy = true;
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
            // Reset to page 1 on fresh load (filters changed or refresh clicked)
            CurrentPage = 1;

            // Always fetch ALL data to ensure TotalRecords is accurate
            return await Service.GetAllAuditItemsAsync();
        }

        protected override void FilterData()
        {
            if (AllItems == null) return;

            var query = AllItems.AsEnumerable();

            // 1. Apply Text Filter
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

            // 2. Apply AR No Filter
            if (!string.IsNullOrWhiteSpace(SearchArNo))
            {
                var tokens = SearchArNo.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim().ToLower());
                query = query.Where(i =>
                {
                    string dbAr = i.ArNo?.ToLower() ?? "";
                    return tokens.Any(token => dbAr.Contains(token));
                });
            }

            // 3. Apply Status Filter
            if (SelectedStatusOption != null && SelectedStatusOption.Value != "All")
            {
                query = query.Where(i => i.Status.ToString() == SelectedStatusOption.Value);
            }

            // 4. Apply 'My Requests Only' Filter
            if (!ShowAll)
            {
                query = query.Where(i => string.Equals(i.RequestedBy, Username, StringComparison.OrdinalIgnoreCase));
            }

            // 5. Store filtered results and update counts
            _filteredItemsList = query.ToList();
            TotalRecords = _filteredItemsList.Count;

            // 6. Calculate Paging
            IsPagingEnabled = true;
            TotalPages = (int)Math.Ceiling((double)TotalRecords / PageSize);
            if (TotalPages < 1) TotalPages = 1;

            if (CurrentPage > TotalPages) CurrentPage = TotalPages;
            if (CurrentPage < 1) CurrentPage = 1;

            // 7. Apply Paging to View
            var pagedList = _filteredItemsList.Skip((CurrentPage - 1) * PageSize).Take(PageSize).ToList();
            Items = new ObservableCollection<AwrItemQueueDto>(pagedList);
        }
    }
}