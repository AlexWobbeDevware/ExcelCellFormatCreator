using ExcelTemplateCellStyleCreator.Core;
using System.Collections.ObjectModel;
using System.Drawing.Text;
using System.Globalization;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

namespace ExcelTemplateCellStyleCreator.WPF
{
    /// <summary>
    /// Main view model. Owns form state, commands, and delegates to Core services
    /// for validation, export, and file reading.
    /// </summary>
    public class MainViewModel : ViewModelBase
    {
        private readonly StyleDefaults _defaults = new();

        // ─── Form input backing fields ──────────────────────────────────────────

        private string _selectedFontName = string.Empty;
        private string _fontSize = string.Empty;
        private string _fontColor = string.Empty;
        private string _bgColor = string.Empty;
        private bool _isBold;
        private bool _isItalic;
        private bool _borderLeft = true;
        private bool _borderRight = true;
        private bool _borderTop = true;
        private bool _borderBottom = true;
        private bool _enableAlignment;
        private string _selectedHorizontalAlignment = "Left";
        private string _selectedVerticalAlignment = "Center";
        private bool _wrapText;

        // ─── Edit state ─────────────────────────────────────────────────────────

        private StyleRowViewModel? _selectedStyleRow;
        private bool _isEditing;

        // ─── Validation error backing fields ────────────────────────────────────

        private string _fontNameError = string.Empty;
        private string _fontSizeError = string.Empty;
        private string _fontColorError = string.Empty;
        private string _bgColorError = string.Empty;

        // ─── Status backing fields ──────────────────────────────────────────────

        private string _statusText = string.Empty;
        private SolidColorBrush _statusBrush = Brushes.Gray;

        // ─── Constructor ────────────────────────────────────────────────────────

        public MainViewModel()
        {
            using var fonts = new InstalledFontCollection();
            FontNames = fonts.Families.OrderBy(f => f.Name).Select(f => f.Name).ToList();

            LoadDefaults();

            AddStyleCommand      = new RelayCommand(_ => AddOrUpdateStyle());
            DeleteStyleCommand   = new RelayCommand(_ => DeleteStyle());
            CancelEditCommand    = new RelayCommand(_ => CancelEdit());
            GenerateExcelCommand = new RelayCommand(_ => GenerateExcel());
            OpenFileCommand      = new RelayCommand(_ => OpenFile());
        }

        // ─── Collections ────────────────────────────────────────────────────────

        public List<string> FontNames { get; }
        public List<string> HorizontalAlignments { get; } = ["Left", "Center", "Right"];
        public List<string> VerticalAlignments { get; } = ["Top", "Center", "Bottom"];
        public ObservableCollection<StyleRowViewModel> StyleRows { get; } = new();

        // ─── Commands ───────────────────────────────────────────────────────────

        public ICommand AddStyleCommand { get; }
        public ICommand DeleteStyleCommand { get; }
        public ICommand CancelEditCommand { get; }
        public ICommand GenerateExcelCommand { get; }
        public ICommand OpenFileCommand { get; }

        // ─── Edit state properties ──────────────────────────────────────────────

        public StyleRowViewModel? SelectedStyleRow
        {
            get => _selectedStyleRow;
            set
            {
                if (SetProperty(ref _selectedStyleRow, value))
                {
                    if (value != null) LoadRowIntoForm(value);
                    IsEditing = value != null;
                }
            }
        }

        public bool IsEditing
        {
            get => _isEditing;
            private set { if (SetProperty(ref _isEditing, value)) OnPropertyChanged(nameof(AddOrUpdateButtonText)); }
        }

        public string AddOrUpdateButtonText => IsEditing ? "Update Style" : "Add Style";

        // ─── Form input properties ──────────────────────────────────────────────

        public string SelectedFontName { get => _selectedFontName; set => SetProperty(ref _selectedFontName, value); }
        public string FontSize { get => _fontSize; set => SetProperty(ref _fontSize, value); }
        public string FontColor { get => _fontColor; set => SetProperty(ref _fontColor, value); }
        public string BgColor { get => _bgColor; set => SetProperty(ref _bgColor, value); }

        public bool IsBold
        {
            get => _isBold;
            set { if (SetProperty(ref _isBold, value)) OnPropertyChanged(nameof(PreviewFontWeight)); }
        }

        public bool IsItalic
        {
            get => _isItalic;
            set { if (SetProperty(ref _isItalic, value)) OnPropertyChanged(nameof(PreviewFontStyle)); }
        }

        public bool BorderLeft
        {
            get => _borderLeft;
            set { if (SetProperty(ref _borderLeft, value)) OnPropertyChanged(nameof(PreviewBorderThickness)); }
        }

        public bool BorderRight
        {
            get => _borderRight;
            set { if (SetProperty(ref _borderRight, value)) OnPropertyChanged(nameof(PreviewBorderThickness)); }
        }

        public bool BorderTop
        {
            get => _borderTop;
            set { if (SetProperty(ref _borderTop, value)) OnPropertyChanged(nameof(PreviewBorderThickness)); }
        }

        public bool BorderBottom
        {
            get => _borderBottom;
            set { if (SetProperty(ref _borderBottom, value)) OnPropertyChanged(nameof(PreviewBorderThickness)); }
        }

        public bool EnableAlignment { get => _enableAlignment; set => SetProperty(ref _enableAlignment, value); }

        public string SelectedHorizontalAlignment
        {
            get => _selectedHorizontalAlignment;
            set { if (SetProperty(ref _selectedHorizontalAlignment, value)) OnPropertyChanged(nameof(PreviewHAlign)); }
        }

        public string SelectedVerticalAlignment
        {
            get => _selectedVerticalAlignment;
            set { if (SetProperty(ref _selectedVerticalAlignment, value)) OnPropertyChanged(nameof(PreviewVAlign)); }
        }

        public bool WrapText
        {
            get => _wrapText;
            set { if (SetProperty(ref _wrapText, value)) OnPropertyChanged(nameof(PreviewTextWrapping)); }
        }

        // ─── Preview properties ─────────────────────────────────────────────────

        public Thickness PreviewBorderThickness => new(
            BorderLeft ? 1 : 0, BorderTop ? 1 : 0,
            BorderRight ? 1 : 0, BorderBottom ? 1 : 0);

        public FontWeight PreviewFontWeight => IsBold ? FontWeights.Bold : FontWeights.Normal;
        public FontStyle PreviewFontStyle => IsItalic ? FontStyles.Italic : FontStyles.Normal;

        public HorizontalAlignment PreviewHAlign => SelectedHorizontalAlignment switch
        {
            "Center" => HorizontalAlignment.Center,
            "Right"  => HorizontalAlignment.Right,
            _        => HorizontalAlignment.Left
        };

        public VerticalAlignment PreviewVAlign => SelectedVerticalAlignment switch
        {
            "Center" => VerticalAlignment.Center,
            "Bottom" => VerticalAlignment.Bottom,
            _        => VerticalAlignment.Top
        };

        public TextWrapping PreviewTextWrapping => WrapText ? TextWrapping.Wrap : TextWrapping.NoWrap;

        // ─── Validation errors ──────────────────────────────────────────────────

        public string FontNameError { get => _fontNameError; set => SetProperty(ref _fontNameError, value); }
        public string FontSizeError { get => _fontSizeError; set => SetProperty(ref _fontSizeError, value); }
        public string FontColorError { get => _fontColorError; set => SetProperty(ref _fontColorError, value); }
        public string BgColorError { get => _bgColorError; set => SetProperty(ref _bgColorError, value); }

        // ─── Status ─────────────────────────────────────────────────────────────

        public string StatusText { get => _statusText; set => SetProperty(ref _statusText, value); }
        public SolidColorBrush StatusBrush { get => _statusBrush; set => SetProperty(ref _statusBrush, value); }

        // ─── Form ↔ Data mapping ────────────────────────────────────────────────

        private void LoadDefaults()
        {
            _selectedFontName = FontNames.Contains(_defaults.FontName) ? _defaults.FontName : FontNames.FirstOrDefault() ?? string.Empty;
            _fontSize  = _defaults.FontSize.ToString(CultureInfo.InvariantCulture);
            _fontColor = _defaults.FontColor;
            _bgColor   = _defaults.BgColor;
            _isBold    = _defaults.IsBold;
            _isItalic  = _defaults.IsItalic;

            var borders = BorderHelper.ParseSelection(_defaults.BorderSelection);
            _borderLeft = borders.Left; _borderRight = borders.Right;
            _borderTop = borders.Top; _borderBottom = borders.Bottom;
        }

        private void ResetFormToDefaults()
        {
            SelectedFontName = FontNames.Contains(_defaults.FontName) ? _defaults.FontName : FontNames.FirstOrDefault() ?? string.Empty;
            FontSize  = _defaults.FontSize.ToString(CultureInfo.InvariantCulture);
            FontColor = _defaults.FontColor;
            BgColor   = _defaults.BgColor;
            IsBold    = _defaults.IsBold;
            IsItalic  = _defaults.IsItalic;

            var borders = BorderHelper.ParseSelection(_defaults.BorderSelection);
            BorderLeft = borders.Left; BorderRight = borders.Right;
            BorderTop = borders.Top; BorderBottom = borders.Bottom;

            EnableAlignment = false;
            SelectedHorizontalAlignment = "Left";
            SelectedVerticalAlignment = "Center";
            WrapText = false;
        }

        private void LoadRowIntoForm(StyleRowViewModel row)
        {
            SelectedFontName = row.FontName;
            FontSize  = row.FontSize.ToString(CultureInfo.InvariantCulture);
            FontColor = row.FontColor;
            BgColor   = row.BgColor;
            IsBold    = row.IsBold;
            IsItalic  = row.IsItalic;

            var borders = BorderHelper.ParseSelection(row.BorderSelection);
            BorderLeft = borders.Left; BorderRight = borders.Right;
            BorderTop = borders.Top; BorderBottom = borders.Bottom;

            EnableAlignment             = row.HorizontalAlign != "Left" || row.VerticalAlign != "Center" || row.WrapText;
            SelectedHorizontalAlignment = row.HorizontalAlign;
            SelectedVerticalAlignment   = row.VerticalAlign;
            WrapText                    = row.WrapText;
        }

        private StyleData BuildStyleDataFromForm(double parsedFontSize)
        {
            string borderSelection = BorderHelper.BuildSelectionString(BorderLeft, BorderRight, BorderTop, BorderBottom);
            return StyleDataMapper.Create(
                SelectedFontName, parsedFontSize, FontColor, BgColor,
                IsBold, IsItalic, borderSelection,
                EnableAlignment, SelectedHorizontalAlignment, SelectedVerticalAlignment, WrapText);
        }

        private static void ApplyStyleDataToRow(StyleRowViewModel row, StyleData data)
        {
            row.FontName        = data.FontName;
            row.FontSize        = data.FontSize;
            row.FontColor       = data.FontColor;
            row.BgColor         = data.BgColor;
            row.IsBold          = data.IsBold;
            row.IsItalic        = data.IsItalic;
            row.BorderSelection = data.BorderSelection;
            row.HorizontalAlign = data.HorizontalAlign;
            row.VerticalAlign   = data.VerticalAlign;
            row.WrapText        = data.WrapText;
        }

        // ─── Validation ─────────────────────────────────────────────────────────

        private bool Validate(out double parsedFontSize)
        {
            bool valid = true;
            parsedFontSize = 0;

            FontNameError = string.IsNullOrEmpty(SelectedFontName) ? "Select a font name." : string.Empty;
            if (FontNameError.Length > 0) valid = false;

            if (!double.TryParse(FontSize, NumberStyles.Any, CultureInfo.InvariantCulture, out parsedFontSize) || parsedFontSize <= 0)
            { FontSizeError = "Enter a positive number."; valid = false; }
            else FontSizeError = string.Empty;

            FontColorError = !HexColorValidator.IsValid(FontColor.Trim()) ? "Enter a 6-digit hex color (e.g. FF0000)." : string.Empty;
            if (FontColorError.Length > 0) valid = false;

            BgColorError = !HexColorValidator.IsValid(BgColor.Trim()) ? "Enter a 6-digit hex color (e.g. FFFFFF)." : string.Empty;
            if (BgColorError.Length > 0) valid = false;

            return valid;
        }

        private bool IsDuplicate(StyleData data, StyleRowViewModel? skipRow = null)
        {
            foreach (var row in StyleRows)
            {
                if (ReferenceEquals(row, skipRow)) continue;
                if (row.FontName == data.FontName && row.FontSize == data.FontSize &&
                    row.FontColor == data.FontColor && row.BgColor == data.BgColor &&
                    row.IsBold == data.IsBold && row.IsItalic == data.IsItalic &&
                    row.BorderSelection == data.BorderSelection &&
                    row.HorizontalAlign == data.HorizontalAlign &&
                    row.VerticalAlign == data.VerticalAlign && row.WrapText == data.WrapText)
                    return true;
            }
            return false;
        }

        // ─── Commands ───────────────────────────────────────────────────────────

        private void AddOrUpdateStyle()
        {
            if (!Validate(out double parsedFontSize))
            {
                SetStatus("Please fix the highlighted errors.", StatusMessage.ErrorBrush);
                return;
            }

            var data = BuildStyleDataFromForm(parsedFontSize);

            if (IsDuplicate(data, IsEditing ? SelectedStyleRow : null))
            {
                SetStatus("Duplicate style — this combination already exists.", StatusMessage.WarningBrush);
                return;
            }

            if (IsEditing && SelectedStyleRow != null)
            {
                ApplyStyleDataToRow(SelectedStyleRow, data);
                SetStatus("Style updated.", StatusMessage.SuccessBrush);
                SelectedStyleRow = null;
                ResetFormToDefaults();
            }
            else
            {
                var newRow = new StyleRowViewModel();
                ApplyStyleDataToRow(newRow, data);
                StyleRows.Add(newRow);
                RenumberStyleIndices();
                SetStatus($"Style added ({StyleRows.Count} total).", StatusMessage.SuccessBrush);
                UpdateDefaults(data);
            }
        }

        private void DeleteStyle()
        {
            if (SelectedStyleRow == null) return;
            StyleRows.Remove(SelectedStyleRow);
            SelectedStyleRow = null;
            ResetFormToDefaults();
            RenumberStyleIndices();
            SetStatus($"Style deleted ({StyleRows.Count} remaining).", StatusMessage.WarningBrush);
        }

        private void CancelEdit()
        {
            SelectedStyleRow = null;
            ResetFormToDefaults();
            StatusText = string.Empty;
        }

        private void GenerateExcel()
        {
            if (StyleRows.Count == 0)
            {
                SetStatus("Add at least one style before generating.", StatusMessage.WarningBrush);
                return;
            }

            try
            {
                var styles = StyleRows.Select(r => new StyleData
                {
                    FontName = r.FontName, FontSize = r.FontSize,
                    FontColor = r.FontColor, BgColor = r.BgColor,
                    IsBold = r.IsBold, IsItalic = r.IsItalic,
                    BorderSelection = r.BorderSelection,
                    HorizontalAlign = r.HorizontalAlign,
                    VerticalAlign = r.VerticalAlign, WrapText = r.WrapText
                }).ToList();

                var indices = ExcelExportService.Export(StyleConstants.DefaultFilePath, styles);
                for (int i = 0; i < StyleRows.Count; i++)
                    StyleRows[i].StyleIndex = indices[i];

                SetStatus($"File created: {StyleConstants.DefaultFilePath}", StatusMessage.SuccessBrush);
            }
            catch (Exception ex)
            {
                SetStatus($"Error: {ex.Message}", StatusMessage.ErrorBrush);
            }
        }

        private void OpenFile()
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                Title  = "Open Excel File"
            };
            if (dialog.ShowDialog() != true) return;

            try
            {
                var styles = StyleReader.ReadStyles(dialog.FileName);
                SelectedStyleRow = null;
                StyleRows.Clear();
                ResetFormToDefaults();

                foreach (var data in styles)
                {
                    var row = new StyleRowViewModel();
                    ApplyStyleDataToRow(row, data);
                    StyleRows.Add(row);
                }

                RenumberStyleIndices();
                SetStatus($"Loaded {StyleRows.Count} style(s) from {System.IO.Path.GetFileName(dialog.FileName)}", StatusMessage.SuccessBrush);
            }
            catch (Exception ex)
            {
                SetStatus($"Error reading file: {ex.Message}", StatusMessage.ErrorBrush);
            }
        }

        // ─── Helpers ────────────────────────────────────────────────────────────

        private void SetStatus(string text, SolidColorBrush brush)
        {
            StatusText = text;
            StatusBrush = brush;
        }

        private void RenumberStyleIndices()
        {
            for (int i = 0; i < StyleRows.Count; i++)
                StyleRows[i].StyleIndex = (uint)(i + 1);
        }

        private void UpdateDefaults(StyleData data)
        {
            _defaults.FontName        = data.FontName;
            _defaults.FontSize        = data.FontSize;
            _defaults.FontColor       = data.FontColor;
            _defaults.BgColor         = data.BgColor;
            _defaults.IsBold          = data.IsBold;
            _defaults.IsItalic        = data.IsItalic;
            _defaults.BorderSelection = data.BorderSelection == "(none)" ? "" : data.BorderSelection;
        }
    }
}
