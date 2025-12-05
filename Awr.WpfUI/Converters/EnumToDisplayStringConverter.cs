using System;
using System.Globalization;
using System.Windows.Data;
using Awr.Core.Enums;

namespace Awr.WpfUI.Converters
{
    public class EnumToDisplayStringConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is AwrItemStatus status)
            {
                switch (status)
                {
                    case AwrItemStatus.PendingIssuance: return "Pending Approval";
                    case AwrItemStatus.Issued: return "Approved (Ready to Print)";
                    case AwrItemStatus.Received: return "Completed (Printed)";
                    case AwrItemStatus.Voided: return "Voided";
                    case AwrItemStatus.RejectedByQa: return "Rejected";
                    case AwrItemStatus.Draft: return "Draft";
                    default: return status.ToString();
                }
            }
            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}