using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace ExcelTemplateCellStyleCreator
{
    public class StyleManager
    {
        public Fonts Fonts { get; private set; }
        public Fills Fills { get; private set; }
        public Borders Borders { get; private set; }
        public CellFormats CellFormats { get; private set; }

        public StyleManager()
        {
            Fonts = CreateDefaultFonts();
            Fills = CreateDefaultFills();
            Borders = CreateDefaultBorders();
            CellFormats = new CellFormats(new CellFormat());
        }

        private Fonts CreateDefaultFonts()
        {
            var fonts = new Fonts(
                new Font(
                    new FontSize() { Val = 11 },
                    new Color() { Theme = 1 },
                    new FontName() { Val = "Calibri" },
                    new FontFamilyNumbering() { Val = 2 },
                    new FontScheme() { Val = FontSchemeValues.Minor }
                )
            );

            return fonts;
        }

        private Fills CreateDefaultFills()
        {
            var fills = new Fills(
                 new Fill(new PatternFill() { PatternType = PatternValues.None }),
                 new Fill(new PatternFill(
                     new ForegroundColor() { Rgb = new HexBinaryValue("FFFFFFFF") },
                     new BackgroundColor() { Indexed = 64 }) // Indexed 64 is the default background color (white)
                 { PatternType = PatternValues.Solid })
             );

            return fills;
        }

        private Borders CreateDefaultBorders()
        {
            return new Borders(
                new Border(
                    new LeftBorder(),
                    new RightBorder(),
                    new TopBorder(),
                    new BottomBorder(),
                    new DiagonalBorder()
                )
            );
        }

        public uint ConfigureFont(string fontName, double fontSize, string fontColor, bool isBold, bool isItalic)
        {
            for (uint i = 0; i < Fonts.ChildElements.Count; i++)
            {
                Font existingFont = (Font)Fonts.ElementAt((int)i);
                if (existingFont.FontSize.Val == fontSize &&
                    existingFont.Color != null && 
                    existingFont.Color.Rgb != null && 
                    existingFont.Color.Rgb.Value == fontColor &&
                    existingFont.FontName.Val == fontName &&
                    existingFont.Bold != null == isBold &&
                    existingFont.Italic != null == isItalic)
                {
                    return i;
                }
            }

            Font font = new Font();
            font.Append(new FontSize() { Val = fontSize });
            font.Append(new Color() { Rgb = new HexBinaryValue() { Value = fontColor } });
            font.Append(new FontName() { Val = fontName });
            if (isBold)
            {
                font.Append(new Bold());
            }
            if (isItalic)
            {
                font.Append(new Italic());
            }
            Fonts.Append(font);
            return (uint)Fonts.ChildElements.Count - 1;
        }

        public uint ConfigureFills(string bgColor)
        {
            for (uint i = 0; i < Fills.ChildElements.Count; i++)
            {
                Fill existingFill = (Fill)Fills.ElementAt((int)i);
                PatternFill patternFill = existingFill.PatternFill;
                if (patternFill != null && 
                    patternFill.ForegroundColor != null &&
                    patternFill.ForegroundColor.Rgb != null &&
                    patternFill.ForegroundColor.Rgb.Value == bgColor &&
                    patternFill.PatternType == PatternValues.Solid)
                {
                    return i;
                }
            }

            var fill = new Fill(
                new PatternFill(
                    new ForegroundColor() { Rgb = new HexBinaryValue(bgColor) },
                    new BackgroundColor() { Rgb = new HexBinaryValue(bgColor) }
                )
                { PatternType = PatternValues.Solid }
            );
            Fills.Append(fill);
            return (uint)Fills.ChildElements.Count - 1;
        }

        public Border ConfigureBorder(string borderSelection)
        {
            Border border = new Border();
            if (borderSelection.Contains("l")) border.Append(new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin });
            if (borderSelection.Contains("r")) border.Append(new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin });
            if (borderSelection.Contains("t")) border.Append(new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin });
            if (borderSelection.Contains("b")) border.Append(new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin });
            return border;
        }

        public uint GetOrCreateBorderId(Border border)
        {
            uint borderId = 0;
            foreach (var b in Borders.Elements<Border>())
            {
                if (b.OuterXml == border.OuterXml)
                {
                    return borderId;
                }
                borderId++;
            }

            Borders.Append(border);
            return borderId;
        }

        public CellFormat CreateCellFormat(uint fontId, uint fillId, uint borderId, bool configureAlignment, HorizontalAlignmentValues horizontalAlignment, VerticalAlignmentValues verticalAlignment, bool wrapText)
        {
            CellFormat cellFormat = new CellFormat()
            {
                FontId = fontId,
                FillId = fillId,
                BorderId = borderId,
                ApplyFont = true,
                ApplyFill = true,
                ApplyBorder = true,
                ApplyAlignment = configureAlignment,
                Alignment = configureAlignment
                    ? new Alignment() { Horizontal = horizontalAlignment, Vertical = verticalAlignment, WrapText = wrapText }
                    : null
            };

            return cellFormat;
        }

        public bool CellFormatExists(CellFormat newCellFormat)
        {
            foreach (CellFormat cellFormat in CellFormats.Elements<CellFormat>())
            {
                if ((cellFormat.FontId?.Value ?? 0) == (newCellFormat.FontId?.Value ?? 0) &&
                    (cellFormat.FillId?.Value ?? 0) == (newCellFormat.FillId?.Value ?? 0) &&
                    (cellFormat.BorderId?.Value ?? 0) == (newCellFormat.BorderId?.Value ?? 0) &&
                    (cellFormat.ApplyFont?.Value ?? false) == (newCellFormat.ApplyFont?.Value ?? false) &&
                    (cellFormat.ApplyFill?.Value ?? false) == (newCellFormat.ApplyFill?.Value ?? false) &&
                    (cellFormat.ApplyBorder?.Value ?? false) == (newCellFormat.ApplyBorder?.Value ?? false) &&
                    (cellFormat.ApplyAlignment?.Value ?? false) == (newCellFormat.ApplyAlignment?.Value ?? false) &&
                    ((cellFormat.Alignment == null && newCellFormat.Alignment == null) ||
                     (cellFormat.Alignment != null && newCellFormat.Alignment != null &&
                      (cellFormat.Alignment.Horizontal?.Value ?? HorizontalAlignmentValues.Left) == (newCellFormat.Alignment.Horizontal?.Value ?? HorizontalAlignmentValues.Left) &&
                      (cellFormat.Alignment.Vertical?.Value ?? VerticalAlignmentValues.Top) == (newCellFormat.Alignment.Vertical?.Value ?? VerticalAlignmentValues.Top) &&
                      (cellFormat.Alignment.WrapText?.Value ?? false) == (newCellFormat.Alignment.WrapText?.Value ?? false))))
                {
                    return true;
                }
            }

            return false;
        }

        public void SaveStylesheet(WorkbookStylesPart stylesPart, Stylesheet stylesheet)
        {
            stylesheet.Fonts = Fonts;
            stylesheet.Fills = Fills;
            stylesheet.Borders = Borders;
            stylesheet.CellFormats = CellFormats;
            stylesPart.Stylesheet = stylesheet;
            stylesPart.Stylesheet.Save();
        }
    }
}
