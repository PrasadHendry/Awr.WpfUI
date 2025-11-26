using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Awr.Core.DTOs;
using Awr.WpfUI.ViewModels.Shared;

namespace Awr.WpfUI.ViewModels
{
    public class AuditTrailViewModel : WorkQueueViewModel
    {
        private bool _showAll = true;
        public bool ShowAll
        {
            get => _showAll;
            set { if (SetProperty(ref _showAll, value)) _ = LoadDataAsync(); }
        }

        public AuditTrailViewModel(string username) : base(username)
        {
        }

        protected override async Task<List<AwrItemQueueDto>> FetchDataInternalAsync()
        {
            // Logic: If "Show All" is checked, get everything. 
            // Otherwise, get only items submitted by the current user.
            if (!ShowAll)
            {
                return await Service.GetMySubmittedItemsAsync(Username);
            }
            else
            {
                return await Service.GetAllAuditItemsAsync();
            }
        }
    }
}