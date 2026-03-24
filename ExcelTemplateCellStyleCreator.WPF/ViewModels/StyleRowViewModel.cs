namespace ExcelTemplateCellStyleCreator.WPF
{
    /// <summary>
    /// Represents one row in the "Added Styles" DataGrid.
    /// Extends <see cref="ViewModelBase"/> so in-place updates (editing) refresh the DataGrid.
    /// </summary>
    public class StyleRowViewModel : ViewModelBase
    {
        private uint   _styleIndex;
        private string _fontName        = string.Empty;
        private double _fontSize;
        private string _fontColor       = string.Empty;
        private string _bgColor         = string.Empty;
        private bool   _isBold;
        private bool   _isItalic;
        private string _borderSelection = string.Empty;
        private string _horizontalAlign = string.Empty;
        private string _verticalAlign   = string.Empty;
        private bool   _wrapText;

        public uint StyleIndex
        {
            get => _styleIndex;
            set => SetProperty(ref _styleIndex, value);
        }

        public string FontName
        {
            get => _fontName;
            set => SetProperty(ref _fontName, value);
        }

        public double FontSize
        {
            get => _fontSize;
            set => SetProperty(ref _fontSize, value);
        }

        public string FontColor
        {
            get => _fontColor;
            set => SetProperty(ref _fontColor, value);
        }

        public string BgColor
        {
            get => _bgColor;
            set => SetProperty(ref _bgColor, value);
        }

        public bool IsBold
        {
            get => _isBold;
            set => SetProperty(ref _isBold, value);
        }

        public bool IsItalic
        {
            get => _isItalic;
            set => SetProperty(ref _isItalic, value);
        }

        public string BorderSelection
        {
            get => _borderSelection;
            set => SetProperty(ref _borderSelection, value);
        }

        public string HorizontalAlign
        {
            get => _horizontalAlign;
            set => SetProperty(ref _horizontalAlign, value);
        }

        public string VerticalAlign
        {
            get => _verticalAlign;
            set => SetProperty(ref _verticalAlign, value);
        }

        public bool WrapText
        {
            get => _wrapText;
            set => SetProperty(ref _wrapText, value);
        }
    }
}
