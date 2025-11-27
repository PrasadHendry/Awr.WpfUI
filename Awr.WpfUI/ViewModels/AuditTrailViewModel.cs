using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using Awr.Core.DTOs;
using Awr.Core.Enums;
using Awr.WpfUI.ViewModels.Shared;

namespace Awr.WpfUI.ViewModels
{
    public class AuditTrailViewModel : WorkQueueViewModel
    {
        // --- Filter Sources ---
        public List<string> AwrTypes { get; }
        public List<string> Statuses { get; }

        // --- Filter Selections ---
        private string _selectedType;
        public string SelectedType
        {
            get => _selectedType;
            set { if (SetProperty(ref _selectedType, value)) FilterData(); }
        }

        private string _selectedStatus;
        public string SelectedStatus
        {
            get => _selectedStatus;
            set { if (SetProperty(ref _selectedStatus, value)) FilterData(); }
        }

        private bool _showAll = true;
        public bool ShowAll
        {
            get => _showAll;
            set { if (SetProperty(ref _showAll, value)) FilterData(); }
        }

        public AuditTrailViewModel(string username) : base(username)
        {
            // Initialize Filter Lists
            AwrTypes = new List<string> { "--- All Types ---" };
            AwrTypes.AddRange(Enum.GetNames(typeof(AwrType)));

            // FIX: Exclude Draft and Complete from the UI Dropdown
            Statuses = new List<string> { "--- All Statuses ---" };
            foreach (var statusName in Enum.GetNames(typeof(AwrItemStatus)))
            {
                if (statusName != "Draft" && statusName != "Complete")
                {
                    Statuses.Add(statusName);
                }
            }

            // Defaults
            _selectedType = AwrTypes[0];
            _selectedStatus = Statuses[0];
        }

        // Fetch ALL data from DB once, then filter in memory for speed
        protected override async Task<List<AwrItemQueueDto>> FetchDataInternalAsync()
        {
            return await Service.GetAllAuditItemsAsync();
        }

        // Override the Base FilterData to include specific dropdown logic
        protected override void FilterData()
        {
            if (AllItems == null) return;

            var query = AllItems.AsEnumerable();

            // 1. Text Search
            if (!string.IsNullOrWhiteSpace(SearchText))
            {
                string lower = SearchText.ToLower();
                query = query.Where(i =>
                    (i.RequestNo?.ToLower().Contains(lower) ?? false) ||
                    (i.AwrNo?.ToLower().Contains(lower) ?? false) ||
                    (i.MaterialProduct?.ToLower().Contains(lower) ?? false) ||
                    (i.BatchNo?.ToLower().Contains(lower) ?? false)
                );
            }

            // 2. Type Filter
            if (SelectedType != "--- All Types ---")
            {
                query = query.Where(i => i.AwrType.ToString() == SelectedType);
            }

            // 3. Status Filter
            if (SelectedStatus != "--- All Statuses ---")
            {
                query = query.Where(i => i.Status.ToString() == SelectedStatus);
            }

            // 4. "My Requests Only"
            if (!ShowAll)
            {
                query = query.Where(i => string.Equals(i.RequestedBy, Username, StringComparison.OrdinalIgnoreCase));
            }

            Items = new ObservableCollection<AwrItemQueueDto>(query);
        }
    }
}