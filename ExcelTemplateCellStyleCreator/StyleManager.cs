using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelTemplateCellStyleCreator
{
    /// <summary>
    /// Encapsulates text alignment settings for a cell format.
    /// When <see cref="Enabled"/> is <see langword="false"/> the alignment properties are
    /// ignored and no <c>&lt;alignment&gt;</c> element is written to the stylesheet.
    /// </summary>
    /// <param name="Enabled">Whether alignment settings should be applied to the cell format.</param>
    /// <param name="Horizontal">Horizontal text alignment (left, center, right).</param>
    /// <param name="Vertical">Vertical text alignment (top, center, bottom).</param>
    /// <param name="WrapText">Whether text wrapping is enabled.</param>
    public record AlignmentConfig(
        bool Enabled,
        HorizontalAlignmentValues Horizontal,
        VerticalAlignmentValues Vertical,
        bool WrapText);

    /// <summary>
    /// Manages the OpenXML stylesheet collections (fonts, fills, borders, cell formats)
    /// for the generated workbook. Provides get-or-create helpers that deduplicate entries
    /// so the stylesheet stays as compact as possible.
    /// </summary>
    public class StyleManager
    {
        /// <summary>Gets the fonts collection for the workbook stylesheet.</summary>
        public Fonts Fonts { get; private set; }

        /// <summary>Gets the fills collection for the workbook stylesheet.</summary>
        public Fills Fills { get; private set; }

        /// <summary>Gets the borders collection for the workbook stylesheet.</summary>
        public Borders Borders { get; private set; }

        /// <summary>Gets the cell formats collection for the workbook stylesheet.</summary>
        public CellFormats CellFormats { get; private set; }

        /// <summary>
        /// Initializes a new <see cref="StyleManager"/> with the mandatory default entries
        /// that the Open XML specification requires in every valid stylesheet.
        /// </summary>
        public StyleManager()
        {
            Fonts = CreateDefaultFonts();
            Fills = CreateDefaultFills();
            Borders = CreateDefaultBorders();
            CellFormats = new CellFormats(new CellFormat());
        }

        private Fonts CreateDefaultFonts()
        {
            return new Fonts(
                new Font(
                    new FontSize() { Val = 11 },
                    new Color() { Theme = 1 },
                    new FontName() { Val = "Calibri" },
                    new FontFamilyNumbering() { Val = 2 },
                    new FontScheme() { Val = FontSchemeValues.Minor }
                )
            );
        }

        private Fills CreateDefaultFills()
        {
            return new Fills(
                new Fill(new PatternFill() { PatternType = PatternValues.None }),
                new Fill(new PatternFill(
                    new ForegroundColor() { Rgb = new HexBinaryValue("FFFFFFFF") },
                    new BackgroundColor() { Indexed = 64 })
                { PatternType = PatternValues.Solid })
            );
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

        /// <summary>
        /// Returns the index of an existing font that matches all given properties,
        /// or appends a new font and returns its index.
        /// </summary>
        /// <param name="fontName">Font family name (must be installed on the system).</param>
        /// <param name="fontSize">Font size in points.</param>
        /// <param name="fontColor">Font color as a 6-digit hex RGB string.</param>
        /// <param name="isBold">Whether the font is bold.</param>
        /// <param name="isItalic">Whether the font is italic.</param>
        /// <returns>Zero-based index of the font in <see cref="Fonts"/>.</returns>
        public uint GetOrCreateFontId(string fontName, double fontSize, string fontColor, bool isBold, bool isItalic)
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
                font.Append(new Bold());
            if (isItalic)
                font.Append(new Italic());

            Fonts.Append(font);
            return (uint)Fonts.ChildElements.Count - 1;
        }

        /// <summary>
        /// Returns the index of an existing solid fill that matches <paramref name="bgColor"/>,
        /// or appends a new fill and returns its index.
        /// </summary>
        /// <param name="bgColor">Background color as a 6-digit hex RGB string.</param>
        /// <returns>Zero-based index of the fill in <see cref="Fills"/>.</returns>
        public uint GetOrCreateFillId(string bgColor)
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

        /// <summary>
        /// Builds a <see cref="Border"/> element from a selection string.
        /// Each character in <paramref name="borderSelection"/> activates the corresponding side
        /// with a thin auto-color border: 'l' = left, 'r' = right, 't' = top, 'b' = bottom.
        /// </summary>
        /// <param name="borderSelection">String containing any combination of 'l', 'r', 't', 'b'.</param>
        /// <returns>A new <see cref="Border"/> element (not yet registered in <see cref="Borders"/>).</returns>
        public Border CreateBorder(string borderSelection)
        {
            Border border = new Border();
            if (borderSelection.Contains("l")) border.Append(new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin });
            if (borderSelection.Contains("r")) border.Append(new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin });
            if (borderSelection.Contains("t")) border.Append(new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin });
            if (borderSelection.Contains("b")) border.Append(new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin });
            return border;
        }

        /// <summary>
        /// Returns the index of an existing border whose serialized XML matches
        /// <paramref name="border"/>, or appends it and returns its new index.
        /// OuterXml comparison is used because <see cref="Border"/> has no structural equality.
        /// </summary>
        /// <param name="border">Border element to look up or register.</param>
        /// <returns>Zero-based index of the border in <see cref="Borders"/>.</returns>
        public uint GetOrCreateBorderId(Border border)
        {
            uint borderId = 0;
            foreach (var b in Borders.Elements<Border>())
            {
                // Compare serialized XML since Border has no value equality.
                if (b.OuterXml == border.OuterXml)
                    return borderId;
                borderId++;
            }

            Borders.Append(border);
            return borderId;
        }

        /// <summary>
        /// Builds a <see cref="CellFormat"/> that combines the given font, fill, and border
        /// with optional alignment settings.
        /// </summary>
        /// <param name="fontId">Index of the font in <see cref="Fonts"/>.</param>
        /// <param name="fillId">Index of the fill in <see cref="Fills"/>.</param>
        /// <param name="borderId">Index of the border in <see cref="Borders"/>.</param>
        /// <param name="alignment">Alignment configuration; no alignment element is written when <see cref="AlignmentConfig.Enabled"/> is <see langword="false"/>.</param>
        /// <returns>A new <see cref="CellFormat"/> ready to be appended to <see cref="CellFormats"/>.</returns>
        public CellFormat CreateCellFormat(uint fontId, uint fillId, uint borderId, AlignmentConfig alignment)
        {
            return new CellFormat()
            {
                FontId = fontId,
                FillId = fillId,
                BorderId = borderId,
                ApplyFont = true,
                ApplyFill = true,
                ApplyBorder = true,
                ApplyAlignment = alignment.Enabled,
                Alignment = alignment.Enabled
                    ? new Alignment() { Horizontal = alignment.Horizontal, Vertical = alignment.Vertical, WrapText = alignment.WrapText }
                    : null
            };
        }

        /// <summary>
        /// Returns <see langword="true"/> if an identical cell format already exists in
        /// <see cref="CellFormats"/>, preventing duplicate stylesheet entries.
        /// </summary>
        /// <param name="newFormat">The cell format to check for duplicates.</param>
        /// <returns><see langword="true"/> if a matching format is found; otherwise <see langword="false"/>.</returns>
        public bool CellFormatExists(CellFormat newFormat)
        {
            foreach (CellFormat existing in CellFormats.Elements<CellFormat>())
            {
                if (BasicPropertiesMatch(existing, newFormat) && AlignmentsMatch(existing.Alignment, newFormat.Alignment))
                    return true;
            }
            return false;
        }

        /// <summary>
        /// Compares the font, fill, border IDs and apply-flags of two cell formats.
        /// </summary>
        private static bool BasicPropertiesMatch(CellFormat a, CellFormat b)
        {
            return (a.FontId?.Value ?? 0) == (b.FontId?.Value ?? 0) &&
                   (a.FillId?.Value ?? 0) == (b.FillId?.Value ?? 0) &&
                   (a.BorderId?.Value ?? 0) == (b.BorderId?.Value ?? 0) &&
                   (a.ApplyFont?.Value ?? false) == (b.ApplyFont?.Value ?? false) &&
                   (a.ApplyFill?.Value ?? false) == (b.ApplyFill?.Value ?? false) &&
                   (a.ApplyBorder?.Value ?? false) == (b.ApplyBorder?.Value ?? false) &&
                   (a.ApplyAlignment?.Value ?? false) == (b.ApplyAlignment?.Value ?? false);
        }

        /// <summary>
        /// Compares two nullable <see cref="Alignment"/> elements for logical equality,
        /// treating a null element as "default alignment" (left, top, no wrap).
        /// </summary>
        private static bool AlignmentsMatch(Alignment a, Alignment b)
        {
            if (a == null && b == null) return true;
            if (a == null || b == null) return false;

            return (a.Horizontal?.Value ?? HorizontalAlignmentValues.Left) == (b.Horizontal?.Value ?? HorizontalAlignmentValues.Left) &&
                   (a.Vertical?.Value ?? VerticalAlignmentValues.Top) == (b.Vertical?.Value ?? VerticalAlignmentValues.Top) &&
                   (a.WrapText?.Value ?? false) == (b.WrapText?.Value ?? false);
        }

        /// <summary>
        /// Writes all collected fonts, fills, borders, and cell formats into
        /// <paramref name="stylesheet"/> and saves it to <paramref name="stylesPart"/>.
        /// Must be called once after all styles have been defined.
        /// </summary>
        /// <param name="stylesPart">The workbook styles part to persist the stylesheet to.</param>
        /// <param name="stylesheet">The stylesheet element to populate.</param>
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
