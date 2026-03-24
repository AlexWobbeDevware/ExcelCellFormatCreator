namespace ExcelTemplateCellStyleCreator.Core
{
    /// <summary>
    /// Constructs <see cref="StyleData"/> instances from individual style property values.
    /// </summary>
    public static class StyleDataMapper
    {
        /// <summary>
        /// Creates a <see cref="StyleData"/> from explicit property values.
        /// Colors are normalized to uppercase 6-digit hex. Border selection uses "(none)" for display if empty.
        /// </summary>
        public static StyleData Create(
            string fontName, double fontSize, string fontColor, string bgColor,
            bool isBold, bool isItalic, string borderSelection,
            bool enableAlignment, string horizontalAlign, string verticalAlign, bool wrapText)
        {
            return new StyleData
            {
                FontName        = fontName,
                FontSize        = fontSize,
                FontColor       = HexColorValidator.Normalize(fontColor),
                BgColor         = HexColorValidator.Normalize(bgColor),
                IsBold          = isBold,
                IsItalic        = isItalic,
                BorderSelection = BorderHelper.FormatForDisplay(borderSelection),
                HorizontalAlign = enableAlignment ? horizontalAlign : "Left",
                VerticalAlign   = enableAlignment ? verticalAlign : "Center",
                WrapText        = enableAlignment && wrapText
            };
        }
    }
}
