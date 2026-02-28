using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace MIR.Converters
{
    public class BoolToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            bool bValue = false;
            
            if (value is bool b)
                bValue = b;
            else if (value is string s)
                bValue = !string.IsNullOrEmpty(s);

            if (parameter?.ToString() == "Invert")
                bValue = !bValue;

            return bValue ? Visibility.Visible : Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
