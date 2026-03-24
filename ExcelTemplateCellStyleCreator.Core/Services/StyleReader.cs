using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelTemplateCellStyleCreator.Core
{
    /// <summary>
    /// Reads cell styles from an existing Excel file's stylesheet.
    /// </summary>
    public static class StyleReader
    {
        /// <summary>
        /// Opens an Excel file read-only and extracts all non-default cell styles.
        /// </summary>
        public static List<StyleData> ReadStyles(string filePath)
        {
            var results = new List<StyleData>();

            using var document = SpreadsheetDocument.Open(filePath, false);
            var stylesPart = document.WorkbookPart?.WorkbookStylesPart;
            if (stylesPart?.Stylesheet == null) return results;

            var stylesheet = stylesPart.Stylesheet;
            var fonts   = stylesheet.Fonts;
            var fills   = stylesheet.Fills;
            var borders = stylesheet.Borders;
            var formats = stylesheet.CellFormats;

            if (formats == null) return results;

            int index = 0;
            foreach (CellFormat cf in formats.Elements<CellFormat>())
            {
                if (index++ == 0) continue;

                var data = new StyleData();

                ResolveFont(data, fonts, cf);
                ResolveFill(data, fills, cf);
                ResolveBorder(data, borders, cf);
                ResolveAlignment(data, cf);

                results.Add(data);
            }

            return results;
        }

        private static void ResolveFont(StyleData data, Fonts? fonts, CellFormat cf)
        {
            if (fonts == null || cf.FontId?.Value is not uint fontId || fontId >= fonts.ChildElements.Count)
                return;

            var font = (Font)fonts.ElementAt((int)fontId);
            data.FontName  = font.FontName?.Val?.Value ?? StyleConstants.DefaultFontName;
            data.FontSize  = font.FontSize?.Val?.Value ?? StyleConstants.DefaultFontSize;
            data.FontColor = ExtractFontColor(font);
            data.IsBold    = font.Bold != null;
            data.IsItalic  = font.Italic != null;
        }

        private static void ResolveFill(StyleData data, Fills? fills, CellFormat cf)
        {
            if (fills == null || cf.FillId?.Value is not uint fillId || fillId >= fills.ChildElements.Count)
                return;

            var fill = (Fill)fills.ElementAt((int)fillId);
            data.BgColor = ExtractFillColor(fill);
        }

        private static void ResolveBorder(StyleData data, Borders? borders, CellFormat cf)
        {
            if (borders == null || cf.BorderId?.Value is not uint borderId || borderId >= borders.ChildElements.Count)
                return;

            var border = (Border)borders.ElementAt((int)borderId);
            data.BorderSelection = ExtractBorderSelection(border);
        }

        private static void ResolveAlignment(StyleData data, CellFormat cf)
        {
            if (cf.ApplyAlignment?.Value != true || cf.Alignment == null)
                return;

            data.HorizontalAlign = AlignmentMapper.ToHorizontalLabel(cf.Alignment.Horizontal?.Value);
            data.VerticalAlign   = AlignmentMapper.ToVerticalLabel(cf.Alignment.Vertical?.Value);
            data.WrapText        = cf.Alignment.WrapText?.Value ?? false;
        }

        private static string ExtractFontColor(Font font)
        {
            if (font.Color?.Rgb?.Value is string rgb)
                return HexColorValidator.Normalize(rgb);

            if (font.Color?.Theme?.Value is uint theme)
                return theme == 0 ? StyleConstants.DefaultBgColor : StyleConstants.DefaultFontColor;

            return StyleConstants.DefaultFontColor;
        }

        private static string ExtractFillColor(Fill fill)
        {
            var pf = fill.PatternFill;
            if (pf == null || pf.PatternType?.Value != PatternValues.Solid)
                return StyleConstants.DefaultBgColor;

            if (pf.ForegroundColor?.Rgb?.Value is string rgb)
                return HexColorValidator.Normalize(rgb);

            return StyleConstants.DefaultBgColor;
        }

        private static string ExtractBorderSelection(Border border)
        {
            var sb = new System.Text.StringBuilder(4);
            if (HasBorderStyle(border.LeftBorder))   sb.Append('l');
            if (HasBorderStyle(border.RightBorder))  sb.Append('r');
            if (HasBorderStyle(border.TopBorder))    sb.Append('t');
            if (HasBorderStyle(border.BottomBorder)) sb.Append('b');
            return sb.ToString();
        }

        private static bool HasBorderStyle(BorderPropertiesType? bp)
            => bp?.Style != null && bp.Style.Value != BorderStyleValues.None;
    }
}
