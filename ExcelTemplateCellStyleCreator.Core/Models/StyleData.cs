namespace ExcelTemplateCellStyleCreator.Core
{
    /// <summary>
    /// Plain data transfer object representing one cell style.
    /// Used by <see cref="StyleReader"/> and <see cref="ExcelExportService"/>.
    /// </summary>
    public class StyleData
    {
        public string FontName        { get; set; } = StyleConstants.DefaultFontName;
        public double FontSize        { get; set; } = StyleConstants.DefaultFontSize;
        public string FontColor       { get; set; } = StyleConstants.DefaultFontColor;
        public bool   IsBold          { get; set; }
        public bool   IsItalic        { get; set; }
        public string BgColor         { get; set; } = StyleConstants.DefaultBgColor;
        public string BorderSelection { get; set; } = string.Empty;
        public string HorizontalAlign { get; set; } = "Left";
        public string VerticalAlign   { get; set; } = "Center";
        public bool   WrapText        { get; set; }
    }
}
