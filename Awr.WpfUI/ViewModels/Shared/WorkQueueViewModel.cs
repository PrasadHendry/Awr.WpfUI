using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using Awr.Core.DTOs;
using Awr.WpfUI.MvvmCore;
using Awr.WpfUI.Services.Implementation;
using Awr.WpfUI.Services.Interfaces;

namespace Awr.WpfUI.ViewModels.Shared
{
    public abstract class WorkQueueViewModel : BaseViewModel
    {
        protected readonly IWorkflowService Service;
        protected readonly string Username;

        // --- Data ---
        private ObservableCollection<AwrItemQueueDto> _allItems; // Full list from DB
        private ObservableCollection<AwrItemQueueDto> _filteredItems; // Display list
        public ObservableCollection<AwrItemQueueDto> Items { get => _filteredItems; set => SetProperty(ref _filteredItems, value); }

        private AwrItemQueueDto _selectedItem;
        public AwrItemQueueDto SelectedItem
        {
            get => _selectedItem;
            set
            {
                if (SetProperty(ref _selectedItem, value)) OnSelectionChanged();
            }
        }

        // --- State ---
        private bool _isBusy;
        public bool IsBusy { get => _isBusy; set => SetProperty(ref _isBusy, value); }

        private string _statusMessage;
        public string StatusMessage { get => _statusMessage; set => SetProperty(ref _statusMessage, value); }

        private string _searchText;
        public string SearchText
        {
            get => _searchText;
            set
            {
                if (SetProperty(ref _searchText, value)) FilterData();
            }
        }

        public ICommand RefreshCommand { get; }

        protected WorkQueueViewModel(string username)
        {
            Username = username;
            Service = new WorkflowService(); // Manual DI
            RefreshCommand = new RelayCommand(async _ => await LoadDataAsync());
            Items = new ObservableCollection<AwrItemQueueDto>();
            // Auto-load on creation
            _ = LoadDataAsync();
        }

        public async Task LoadDataAsync()
        {
            IsBusy = true;
            StatusMessage = "Loading...";
            try
            {
                var data = await FetchDataInternalAsync();
                _allItems = new ObservableCollection<AwrItemQueueDto>(data);
                FilterData(); // Applies search text if any
                StatusMessage = $"{Items.Count} Records Found.";
            }
            catch (Exception ex)
            {
                StatusMessage = "Error loading data.";
                MessageBox.Show(ex.Message, "Error");
            }
            finally
            {
                IsBusy = false;
            }
        }

        private void FilterData()
        {
            if (_allItems == null) return;

            if (string.IsNullOrWhiteSpace(SearchText))
            {
                Items = new ObservableCollection<AwrItemQueueDto>(_allItems);
            }
            else
            {
                string lower = SearchText.ToLower();
                var filtered = _allItems.Where(i =>
                    (i.RequestNo?.ToLower().Contains(lower) ?? false) ||
                    (i.AwrNo?.ToLower().Contains(lower) ?? false) ||
                    (i.MaterialProduct?.ToLower().Contains(lower) ?? false) ||
                    (i.BatchNo?.ToLower().Contains(lower) ?? false)
                );
                Items = new ObservableCollection<AwrItemQueueDto>(filtered);
            }
        }

        // Abstract method: Each Child VM must define HOW it gets data (e.g., GetIssuanceQueue vs GetReceiptQueue)
        protected abstract Task<List<AwrItemQueueDto>> FetchDataInternalAsync();

        // Virtual method: Optional override for when grid selection changes
        protected virtual void OnSelectionChanged() { }
    }
}