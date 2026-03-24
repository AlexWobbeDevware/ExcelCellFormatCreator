using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelTemplateCellStyleCreator.Core
{
    /// <summary>
    /// Generates an Excel style template file from a list of <see cref="StyleData"/> entries.
    /// Each style is written as a sample row in a "Styles" worksheet.
    /// </summary>
    public static class ExcelExportService
    {
        /// <summary>
        /// Creates an Excel file at <paramref name="filePath"/> containing all given styles.
        /// Deletes any existing file at that path first.
        /// Returns the list of assigned StyleIndex values (parallel to the input list).
        /// </summary>
        public static List<uint> Export(string filePath, IReadOnlyList<StyleData> styles)
        {
            FileManager.DeleteFileIfExists(filePath);

            var styleIndices = new List<uint>(styles.Count);

            using var document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);

            WorkbookPart workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
            sheets.Append(new Sheet
            {
                Id      = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name    = StyleConstants.SheetName
            });

            Worksheet worksheet = new Worksheet();
            HideGridLines(worksheet);

            SheetData sheetData = new SheetData();
            worksheet.Append(sheetData);
            worksheetPart.Worksheet = worksheet;

            WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            Stylesheet stylesheet = new Stylesheet();

            // Rebuild StyleManager from the style list.
            var sm = new StyleManager();
            foreach (var style in styles)
            {
                uint fontId   = sm.GetOrCreateFontId(style.FontName, style.FontSize, style.FontColor, style.IsBold, style.IsItalic);
                uint fillId   = sm.GetOrCreateFillId(style.BgColor);
                string borders = style.BorderSelection == "(none)" ? "" : style.BorderSelection;
                var border     = sm.CreateBorder(borders);
                uint borderId = sm.GetOrCreateBorderId(border);

                bool alignEnabled = style.HorizontalAlign != "Left" || style.VerticalAlign != "Center" || style.WrapText;
                var alignment = new AlignmentConfig(
                    alignEnabled,
                    AlignmentMapper.ToHorizontalEnum(style.HorizontalAlign),
                    AlignmentMapper.ToVerticalEnum(style.VerticalAlign),
                    style.WrapText);

                var cf = sm.CreateCellFormat(fontId, fillId, borderId, alignment);
                sm.CellFormats.Append(cf);

                uint styleIndex = (uint)sm.CellFormats.ChildElements.Count - 1;
                styleIndices.Add(styleIndex);
            }

            // Write sample rows.
            uint rowIndex = 0;
            for (int i = 0; i < styles.Count; i++)
            {
                rowIndex += StyleConstants.RowSpacing;
                var sheetRow = new Row { RowIndex = rowIndex };
                sheetData.Append(sheetRow);

                sheetRow.Append(new Cell
                {
                    CellReference = "A" + rowIndex,
                    CellValue     = new CellValue($"StyleIndex Id = {styleIndices[i]}"),
                    DataType      = CellValues.String
                });
                sheetRow.Append(new Cell
                {
                    CellReference = "B" + rowIndex,
                    CellValue     = new CellValue("Sample Text"),
                    DataType      = CellValues.String,
                    StyleIndex    = styleIndices[i]
                });
            }

            sm.SaveStylesheet(stylesPart, stylesheet);
            worksheetPart.Worksheet.Save();
            workbookPart.Workbook.Save();

            return styleIndices;
        }

        private static void HideGridLines(Worksheet worksheet)
        {
            var sheetViews = new SheetViews();
            sheetViews.Append(new SheetView { WorkbookViewId = (UInt32Value)0U, ShowGridLines = false });
            worksheet.Append(sheetViews);
        }
    }
}
