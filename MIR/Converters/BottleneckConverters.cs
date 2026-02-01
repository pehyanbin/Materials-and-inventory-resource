using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace MIR.Converters
{
    public class BoolToErrorBrushConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool isBottleneck && isBottleneck)
            {
                return new SolidColorBrush(Color.FromRgb(254, 242, 242)); // Error Light
            }
            return Brushes.Transparent;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => throw new NotImplementedException();
    }

    public class BoolToErrorTextBrushConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool isBottleneck && isBottleneck)
            {
                return new SolidColorBrush(Color.FromRgb(153, 27, 27)); // Error Dark
            }
            return new SolidColorBrush(Color.FromRgb(31, 41, 55)); // Text Dark
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => throw new NotImplementedException();
    }
}
