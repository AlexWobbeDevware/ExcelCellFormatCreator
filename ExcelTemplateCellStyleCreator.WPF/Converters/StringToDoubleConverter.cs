using System.Globalization;
using System.Windows.Data;

namespace ExcelTemplateCellStyleCreator.WPF
{
    /// <summary>
    /// Converts a string to a double for font size binding. Returns 11 (default) on parse failure.
    /// </summary>
    public class StringToDoubleConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string s && double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double d) && d > 0)
                return d;
            return 11.0;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotSupportedException();
    }
}
