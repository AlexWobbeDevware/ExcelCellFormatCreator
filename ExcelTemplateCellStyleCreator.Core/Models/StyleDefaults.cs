using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelTemplateCellStyleCreator.Core
{
    /// <summary>
    /// Carries the last-used style property values so that each interactive prompt
    /// pre-fills sensibly based on the previous style the user created.
    /// </summary>
    public class StyleDefaults
    {
        public string FontName { get; set; } = StyleConstants.DefaultFontName;
        public double FontSize { get; set; } = StyleConstants.DefaultFontSize;
        public string FontColor { get; set; } = StyleConstants.DefaultFontColor;
        public string BgColor { get; set; } = StyleConstants.DefaultBgColor;
        public bool IsBold { get; set; } = false;
        public bool IsItalic { get; set; } = false;
        public string BorderSelection { get; set; } = StyleConstants.DefaultBorderSelection;
        public HorizontalAlignmentValues HorizontalAlignment { get; set; } = HorizontalAlignmentValues.Left;
        public VerticalAlignmentValues VerticalAlignment { get; set; } = VerticalAlignmentValues.Center;
        public bool WrapText { get; set; } = false;
    }
}
