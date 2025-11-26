using System;
using System.Globalization;
using System.Windows.Data;

namespace Awr.WpfUI.Converters
{
    /// <summary>
    /// Converts a boolean to its inverse. Used for toggling checkbox logic.
    /// </summary>
    public class InverseBooleanConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // We only expect boolean values for the source (ViewModel property)
            if (value is bool boolValue)
            {
                return !boolValue;
            }

            // If the value is null or not a bool, return the original value
            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // We only expect boolean values from the target (View control)
            if (value is bool boolValue)
            {
                return !boolValue;
            }
            return value;
        }
    }
}