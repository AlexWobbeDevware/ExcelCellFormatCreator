using System.Windows.Media;

namespace ExcelTemplateCellStyleCreator.WPF
{
    /// <summary>
    /// Provides standard status bar color brushes.
    /// </summary>
    public static class StatusMessage
    {
        public static SolidColorBrush SuccessBrush { get; } = new(Color.FromRgb(0x22, 0x88, 0x22));
        public static SolidColorBrush ErrorBrush { get; } = Brushes.Red;
        public static SolidColorBrush WarningBrush { get; } = Brushes.Orange;
    }
}
