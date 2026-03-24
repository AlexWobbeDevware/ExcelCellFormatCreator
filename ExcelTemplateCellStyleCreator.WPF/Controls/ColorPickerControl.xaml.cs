using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;

namespace ExcelTemplateCellStyleCreator.WPF
{
    /// <summary>
    /// Reusable color picker: hex TextBox + preview swatch + dropdown with common colors.
    /// </summary>
    public partial class ColorPickerControl : UserControl
    {
        public static readonly DependencyProperty HexColorProperty =
            DependencyProperty.Register(
                nameof(HexColor),
                typeof(string),
                typeof(ColorPickerControl),
                new FrameworkPropertyMetadata("000000",
                    FrameworkPropertyMetadataOptions.BindsTwoWayByDefault));

        public string HexColor
        {
            get => (string)GetValue(HexColorProperty);
            set => SetValue(HexColorProperty, value);
        }

        public ColorPickerControl()
        {
            InitializeComponent();
        }

        private void Swatch_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string hex)
            {
                HexColor = hex;

                // Close the popup.
                BtnToggle.IsChecked = false;
            }
        }
    }
}
