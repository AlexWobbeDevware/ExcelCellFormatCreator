using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelTemplateCellStyleCreator;
using System.Globalization;

class Program
{
    static void Main(string[] args)
    {
        string filePath = @"c:\temp\ExcelStyleTemplate.xlsx";
        var culture = CultureInfo.CurrentCulture.TwoLetterISOLanguageName;

        FileManager.DeleteFileIfExists(filePath, culture);

        string lastFontName = "Calibri";
        double lastFontSize = 11;
        string lastFontColor = "000000";
        string lastBgColor = "FFFFFF";
        bool lastIsBold = false;
        bool lastIsItalic = false;
        string lastBorderSelection = "lrtb";
        HorizontalAlignmentValues lastHorizontalAlignment = HorizontalAlignmentValues.Left;
        VerticalAlignmentValues lastVerticalAlignment = VerticalAlignmentValues.Center;
        bool lastWrapText = false;

        try
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Styles" };
                sheets.Append(sheet);

                Worksheet worksheet = new Worksheet();
                HideGridLines(worksheet);

                SheetData sheetData = new SheetData();
                worksheet.Append(sheetData);

                WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                Stylesheet stylesheet = new Stylesheet();

                StyleManager styleManager = new StyleManager();

                bool continueAdding = true;
                uint rowIndex = 2;
                uint columnIndex = 2;

                while (continueAdding)
                {
                    Console.WriteLine(culture == "de" ? "Neuen Stil hinzufügen:" : "Add a new style:");
                    Console.WriteLine(culture == "de" ? "Beispielhafte Farben: Rot: FF0000 | Grün: 00FF00 | Blau: 0000FF | Gelb: FFFF00 | Schwarz: 000000 | Weiß: FFFFFF"
                                                      : "Example colors: Red: FF0000 | Green: 00FF00 | Blue: 0000FF | Yellow: FFFF00 | Black: 000000 | White: FFFFFF");

                    lastFontName = UserInputValidator.ValidateFontName(GetUserInput(culture, "Schriftart", "Font name", lastFontName), culture);
                    lastFontSize = double.Parse(UserInputValidator.ValidateFontSize(GetUserInput(culture, "Schriftgröße", "Font size", lastFontSize.ToString()), culture));
                    lastFontColor = UserInputValidator.ValidateHexColor(GetUserInput(culture, "Schriftfarbe", "Font color", lastFontColor), culture);
                    lastIsBold = UserInputValidator.ValidateYesNoInput(GetUserInput(culture, "Fett (y/n)", "Bold (y/n)", lastIsBold ? "y" : "n"), culture);
                    lastIsItalic = UserInputValidator.ValidateYesNoInput(GetUserInput(culture, "Kursiv (y/n)", "Italic (y/n)", lastIsItalic ? "y" : "n"), culture);
                    lastBgColor = UserInputValidator.ValidateHexColor(GetUserInput(culture, "Hintergrundfarbe", "Background color", lastBgColor), culture);
                    lastBorderSelection = UserInputValidator.ValidateBorderSelection(GetUserInput(culture, "Rahmen auswählen (left, right, top, bottom)", "Select borders (left, right, top, bottom)", lastBorderSelection), culture);

                    uint fontId = styleManager.ConfigureFont(lastFontName, lastFontSize, lastFontColor, lastIsBold, lastIsItalic);
                    uint fillId = styleManager.ConfigureFills(lastBgColor);
                    Border border = styleManager.ConfigureBorder(lastBorderSelection);
                    bool configureAlignment = UserInputValidator.ValidateYesNoInput(GetUserInput(culture, "Textausrichtung und Umbruch konfigurieren", "Configure text alignment and wrapping", "n"), culture);
                    if (configureAlignment)
                    {
                        lastHorizontalAlignment = UserInputValidator.GetHorizontalAlignment(GetUserInput(culture, "Horizontale Ausrichtung (L: Links, C: Zentrum, R: Rechts)", "Horizontal alignment (L: Left, C: Center, R: Right)", lastHorizontalAlignment.ToString().Substring(0, 1)), culture);
                        lastVerticalAlignment = UserInputValidator.GetVerticalAlignment(GetUserInput(culture, "Vertikale Ausrichtung (T: Oben, C: Mitte, B: Unten)", "Vertical alignment (T: Top, C: Center, B: Bottom)", lastVerticalAlignment.ToString().Substring(0, 1)), culture);
                        lastWrapText = UserInputValidator.ValidateYesNoInput(GetUserInput(culture, "Textumbruch aktivieren", "Enable text wrapping", lastWrapText ? "y" : "n"), culture);
                    }
                    uint borderId = styleManager.GetOrCreateBorderId(border);

                    CellFormat cellFormat = styleManager.CreateCellFormat(fontId, fillId, borderId, configureAlignment, lastHorizontalAlignment, lastVerticalAlignment, lastWrapText);

                    if (!styleManager.CellFormatExists(cellFormat))
                    {
                        styleManager.CellFormats.Append(cellFormat);
                    }

                    InsertCellIntoSheet(sheetData, styleManager.CellFormats, ref rowIndex, ref columnIndex);

                    continueAdding = UserInputValidator.ValidateYesNoInput(GetUserInput(culture, "Weiteren Stil hinzufügen", "Add another style (y/n)", "y"), culture);
                }

                styleManager.SaveStylesheet(stylesPart, stylesheet);
                worksheetPart.Worksheet = worksheet;
                worksheetPart.Worksheet.Save();
                workbookPart.Workbook.Save();
            }

            Console.WriteLine(culture == "de" ? $"Excel-Datei wurde erfolgreich erstellt: {filePath}" : $"Excel file successfully created: {filePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine(culture == "de" ? $"Fehler beim Erstellen der Excel-Datei: {ex.Message}" : $"Error creating Excel file: {ex.Message}");
        }
    }

    private static void HideGridLines(Worksheet worksheet)
    {
        SheetViews sheetViews = new SheetViews();
        SheetView sheetView = new SheetView() { WorkbookViewId = (UInt32Value)0U, ShowGridLines = false };
        sheetViews.Append(sheetView);
        worksheet.Append(sheetViews);
    }

    private static string GetUserInput(string culture, string promptDe, string promptEn, string defaultValue)
    {
        Console.WriteLine();
        Console.Write(culture == "de" ? $"{promptDe} (Standard: {defaultValue}): " : $"{promptEn} (Default: {defaultValue}): ");
        string input = Console.ReadLine() ?? string.Empty;
        return string.IsNullOrWhiteSpace(input) ? defaultValue : input;
    }

    private static void InsertCellIntoSheet(SheetData sheetData, CellFormats cellFormats, ref uint rowIndex, ref uint columnIndex)
    {
        // Increment rowIndex to ensure styles are entered in every second row
        rowIndex += 2;

        uint tempRowIndex = rowIndex;

        // Find the row with the specified rowIndex or create a new one
        Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == tempRowIndex);
        if (row == null)
        {
            row = new Row() { RowIndex = rowIndex };
            sheetData.Append(row);
        }

        // Create a new cell in column A with the StyleIndex Id text without styling
        Cell cellA = new Cell()
        {
            CellReference = "A" + rowIndex,
            CellValue = new CellValue($"StyleIndex Id = {cellFormats.ChildElements.Count - 1}"),
            DataType = CellValues.String
        };
        row.Append(cellA);

        // Create a new cell in column B with the specified style
        uint styleIndex = (uint)cellFormats.ChildElements.Count - 1;
        Cell cellB = new Cell()
        {
            CellReference = "B" + rowIndex,
            CellValue = new CellValue("Sample Text"),
            DataType = CellValues.String,
            StyleIndex = styleIndex
        };
        row.Append(cellB);
    }
}
