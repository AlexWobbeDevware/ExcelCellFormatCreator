using ExcelTemplateCellStyleCreator.Core;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace ExcelTemplateCellStyleCreator.WPF
{
    /// <summary>
    /// Converts a 6-digit hex color string to a <see cref="SolidColorBrush"/>.
    /// Returns <see cref="Brushes.Transparent"/> when the input is not a valid hex color.
    /// </summary>
    public class HexToBrushConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string hex && HexColorValidator.IsValid(hex))
            {
                byte r = System.Convert.ToByte(hex[0..2], 16);
                byte g = System.Convert.ToByte(hex[2..4], 16);
                byte b = System.Convert.ToByte(hex[4..6], 16);
                return new SolidColorBrush(Color.FromRgb(r, g, b));
            }
            return Brushes.Transparent;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotSupportedException();
    }
}
