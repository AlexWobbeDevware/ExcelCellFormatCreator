using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace ExcelTemplateCellStyleCreator.WPF
{
    /// <summary>
    /// Converts a string to <see cref="Visibility"/>:
    /// non-empty → <see cref="Visibility.Visible"/>, empty/null → <see cref="Visibility.Collapsed"/>.
    /// Used to show/hide validation error TextBlocks.
    /// </summary>
    public class StringToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            => string.IsNullOrEmpty(value as string) ? Visibility.Collapsed : Visibility.Visible;

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotSupportedException();
    }
}
