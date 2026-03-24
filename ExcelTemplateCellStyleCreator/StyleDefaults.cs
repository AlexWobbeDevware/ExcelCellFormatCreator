using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelTemplateCellStyleCreator
{
    /// <summary>
    /// Carries the last-used style property values so that each interactive prompt
    /// pre-fills sensibly based on the previous style the user created.
    /// A new instance starts with the application's built-in defaults.
    /// </summary>
    public class StyleDefaults
    {
        /// <summary>Gets or sets the font family name. Default: "Calibri".</summary>
        public string FontName { get; set; } = "Calibri";

        /// <summary>Gets or sets the font size in points. Default: 11.</summary>
        public double FontSize { get; set; } = 11;

        /// <summary>Gets or sets the font color as a 6-digit hex RGB string (e.g. "000000" for black). Default: "000000".</summary>
        public string FontColor { get; set; } = "000000";

        /// <summary>Gets or sets the cell background fill color as a 6-digit hex RGB string (e.g. "FFFFFF" for white). Default: "FFFFFF".</summary>
        public string BgColor { get; set; } = "FFFFFF";

        /// <summary>Gets or sets whether the font is bold. Default: false.</summary>
        public bool IsBold { get; set; } = false;

        /// <summary>Gets or sets whether the font is italic. Default: false.</summary>
        public bool IsItalic { get; set; } = false;

        /// <summary>
        /// Gets or sets the active border sides as a string of characters:
        /// 'l' = left, 'r' = right, 't' = top, 'b' = bottom.
        /// Default: "lrtb" (all four sides).
        /// </summary>
        public string BorderSelection { get; set; } = "lrtb";

        /// <summary>Gets or sets the horizontal text alignment. Default: <see cref="HorizontalAlignmentValues.Left"/>.</summary>
        public HorizontalAlignmentValues HorizontalAlignment { get; set; } = HorizontalAlignmentValues.Left;

        /// <summary>Gets or sets the vertical text alignment. Default: <see cref="VerticalAlignmentValues.Center"/>.</summary>
        public VerticalAlignmentValues VerticalAlignment { get; set; } = VerticalAlignmentValues.Center;

        /// <summary>Gets or sets whether text wrapping is enabled. Default: false.</summary>
        public bool WrapText { get; set; } = false;
    }
}
