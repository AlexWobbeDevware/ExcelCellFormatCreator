using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelTemplateCellStyleCreator;
using ExcelTemplateCellStyleCreator.Core;
using System.Globalization;
using static ExcelTemplateCellStyleCreator.Core.LocalizationHelper;

class Program
{
    static void Main(string[] args)
    {
        string filePath = StyleConstants.DefaultFilePath;
        var culture = CultureInfo.CurrentCulture.TwoLetterISOLanguageName;

        if (!FileManager.DeleteFileIfExists(filePath))
        {
            Console.WriteLine(Localize(culture,
                $"Fehler beim Löschen der Datei: {filePath}",
                $"Error deleting file: {filePath}"));
            return;
        }

        try
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                var (worksheetPart, workbookPart, sheetData, stylesPart, stylesheet, styleManager) = CreateDocument(document);

                var defaults = new StyleDefaults();
                uint rowIndex = 2;

                bool continueAdding = true;
                while (continueAdding)
                {
                    CollectAndApplyStyle(culture, styleManager, sheetData, defaults, ref rowIndex);
                    continueAdding = UserInputValidator.ValidateYesNoInput(
                        GetUserInput(culture, "Weiteren Stil hinzufügen", "Add another style (y/n)", "y"), culture);
                }

                styleManager.SaveStylesheet(stylesPart, stylesheet);
                worksheetPart.Worksheet.Save();
                workbookPart.Workbook.Save();
            }

            Console.WriteLine(Localize(culture,
                $"Excel-Datei wurde erfolgreich erstellt: {filePath}",
                $"Excel file successfully created: {filePath}"));
        }
        catch (Exception ex)
        {
            Console.WriteLine(Localize(culture,
                $"Fehler beim Erstellen der Excel-Datei: {ex.Message}",
                $"Error creating Excel file: {ex.Message}"));
        }
    }

    /// <summary>
    /// Initializes the OpenXML workbook structure: workbook part, a single "Styles" worksheet
    /// with grid lines hidden, a styles part, and a <see cref="StyleManager"/> pre-loaded
    /// with the required default stylesheet entries.
    /// </summary>
    /// <param name="document">The newly created <see cref="SpreadsheetDocument"/> to populate.</param>
    /// <returns>
    /// A tuple of:
    /// <list type="bullet">
    ///   <item><see cref="WorksheetPart"/> — the worksheet part (needed to call <c>Save</c> at the end)</item>
    ///   <item><see cref="WorkbookPart"/> — the workbook part (needed to call <c>Save</c> at the end)</item>
    ///   <item><see cref="SheetData"/> — the sheet data rows are appended to</item>
    ///   <item><see cref="WorkbookStylesPart"/> — the styles part passed to <see cref="StyleManager.SaveStylesheet"/></item>
    ///   <item><see cref="Stylesheet"/> — the stylesheet element passed to <see cref="StyleManager.SaveStylesheet"/></item>
    ///   <item><see cref="StyleManager"/> — manages fonts, fills, borders, and cell formats</item>
    /// </list>
    /// </returns>
    private static (WorksheetPart, WorkbookPart, SheetData, WorkbookStylesPart, Stylesheet, StyleManager) CreateDocument(SpreadsheetDocument document)
    {
        WorkbookPart workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

        Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
        sheets.Append(new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Styles" });

        Worksheet worksheet = new Worksheet();
        HideGridLines(worksheet);

        SheetData sheetData = new SheetData();
        worksheet.Append(sheetData);
        worksheetPart.Worksheet = worksheet;

        WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        Stylesheet stylesheet = new Stylesheet();
        StyleManager styleManager = new StyleManager();

        return (worksheetPart, workbookPart, sheetData, stylesPart, stylesheet, styleManager);
    }

    /// <summary>
    /// Interactively collects all style properties from the user, registers the resulting
    /// font / fill / border / alignment in the <paramref name="styleManager"/>, appends the
    /// cell format if it is not a duplicate, and writes a sample row to <paramref name="sheetData"/>.
    /// Updates <paramref name="defaults"/> with the chosen values so the next iteration
    /// pre-fills with sensible defaults.
    /// </summary>
    /// <param name="culture">Two-letter ISO language name for localized prompts.</param>
    /// <param name="styleManager">The style manager that owns all stylesheet collections.</param>
    /// <param name="sheetData">The worksheet's sheet data to append the sample row to.</param>
    /// <param name="defaults">Carries the last-used values for pre-filling prompts.</param>
    /// <param name="rowIndex">Current row index; incremented by 2 to leave a blank row between samples.</param>
    private static void CollectAndApplyStyle(string culture, StyleManager styleManager, SheetData sheetData, StyleDefaults defaults, ref uint rowIndex)
    {
        Console.WriteLine(Localize(culture, "Neuen Stil hinzufügen:", "Add a new style:"));
        Console.WriteLine(Localize(culture,
            "Beispielhafte Farben: Rot: FF0000 | Grün: 00FF00 | Blau: 0000FF | Gelb: FFFF00 | Schwarz: 000000 | Weiß: FFFFFF",
            "Example colors: Red: FF0000 | Green: 00FF00 | Blue: 0000FF | Yellow: FFFF00 | Black: 000000 | White: FFFFFF"));

        defaults.FontName = UserInputValidator.ValidateFontName(GetUserInput(culture, "Schriftart", "Font name", defaults.FontName), culture);
        defaults.FontSize = double.Parse(UserInputValidator.ValidateFontSize(GetUserInput(culture, "Schriftgröße", "Font size", defaults.FontSize.ToString()), culture));
        defaults.FontColor = UserInputValidator.ValidateHexColor(GetUserInput(culture, "Schriftfarbe", "Font color", defaults.FontColor), culture);
        defaults.IsBold = UserInputValidator.ValidateYesNoInput(GetUserInput(culture, "Fett (y/n)", "Bold (y/n)", defaults.IsBold ? "y" : "n"), culture);
        defaults.IsItalic = UserInputValidator.ValidateYesNoInput(GetUserInput(culture, "Kursiv (y/n)", "Italic (y/n)", defaults.IsItalic ? "y" : "n"), culture);
        defaults.BgColor = UserInputValidator.ValidateHexColor(GetUserInput(culture, "Hintergrundfarbe", "Background color", defaults.BgColor), culture);
        defaults.BorderSelection = UserInputValidator.ValidateBorderSelection(GetUserInput(culture, "Rahmen auswählen (left, right, top, bottom)", "Select borders (left, right, top, bottom)", defaults.BorderSelection), culture);

        uint fontId = styleManager.GetOrCreateFontId(defaults.FontName, defaults.FontSize, defaults.FontColor, defaults.IsBold, defaults.IsItalic);
        uint fillId = styleManager.GetOrCreateFillId(defaults.BgColor);
        Border border = styleManager.CreateBorder(defaults.BorderSelection);
        uint borderId = styleManager.GetOrCreateBorderId(border);

        bool configureAlignment = UserInputValidator.ValidateYesNoInput(
            GetUserInput(culture, "Textausrichtung und Umbruch konfigurieren", "Configure text alignment and wrapping", "n"), culture);

        if (configureAlignment)
        {
            defaults.HorizontalAlignment = UserInputValidator.GetHorizontalAlignment(
                GetUserInput(culture, "Horizontale Ausrichtung (L: Links, C: Zentrum, R: Rechts)", "Horizontal alignment (L: Left, C: Center, R: Right)", defaults.HorizontalAlignment.ToString().Substring(0, 1)), culture);
            defaults.VerticalAlignment = UserInputValidator.GetVerticalAlignment(
                GetUserInput(culture, "Vertikale Ausrichtung (T: Oben, C: Mitte, B: Unten)", "Vertical alignment (T: Top, C: Center, B: Bottom)", defaults.VerticalAlignment.ToString().Substring(0, 1)), culture);
            defaults.WrapText = UserInputValidator.ValidateYesNoInput(
                GetUserInput(culture, "Textumbruch aktivieren", "Enable text wrapping", defaults.WrapText ? "y" : "n"), culture);
        }

        var alignment = new AlignmentConfig(configureAlignment, defaults.HorizontalAlignment, defaults.VerticalAlignment, defaults.WrapText);
        CellFormat cellFormat = styleManager.CreateCellFormat(fontId, fillId, borderId, alignment);

        if (!styleManager.CellFormatExists(cellFormat))
        {
            styleManager.CellFormats.Append(cellFormat);
        }

        InsertCellIntoSheet(sheetData, styleManager.CellFormats, ref rowIndex);
    }

    /// <summary>
    /// Disables grid lines on the worksheet for a cleaner style preview experience.
    /// </summary>
    /// <param name="worksheet">The worksheet to configure.</param>
    private static void HideGridLines(Worksheet worksheet)
    {
        SheetViews sheetViews = new SheetViews();
        SheetView sheetView = new SheetView() { WorkbookViewId = (UInt32Value)0U, ShowGridLines = false };
        sheetViews.Append(sheetView);
        worksheet.Append(sheetViews);
    }

    /// <summary>
    /// Writes a localized prompt to the console and reads user input,
    /// returning <paramref name="defaultValue"/> when the user presses Enter without typing.
    /// </summary>
    /// <param name="culture">Two-letter ISO language name to select the correct prompt.</param>
    /// <param name="promptDe">German prompt text (shown when culture is "de").</param>
    /// <param name="promptEn">English prompt text (shown for all other cultures).</param>
    /// <param name="defaultValue">Value to use when the user provides no input.</param>
    /// <returns>The user's trimmed input, or <paramref name="defaultValue"/> if blank.</returns>
    private static string GetUserInput(string culture, string promptDe, string promptEn, string defaultValue)
    {
        Console.WriteLine();
        Console.Write(Localize(culture,
            $"{promptDe} (Standard: {defaultValue}): ",
            $"{promptEn} (Default: {defaultValue}): "));
        string input = Console.ReadLine() ?? string.Empty;
        return string.IsNullOrWhiteSpace(input) ? defaultValue : input;
    }

    /// <summary>
    /// Appends a sample row to <paramref name="sheetData"/> at the current <paramref name="rowIndex"/>:
    /// <list type="bullet">
    ///   <item>Column A — label showing the StyleIndex ID of the new format.</item>
    ///   <item>Column B — "Sample Text" cell with the new style applied.</item>
    /// </list>
    /// </summary>
    /// <param name="sheetData">The worksheet's sheet data to append the row to.</param>
    /// <param name="cellFormats">The current cell formats collection; used to derive the style index.</param>
    /// <param name="rowIndex">
    /// Row index to write to. Incremented by 2 before writing so that a blank row is left
    /// between consecutive style samples for visual clarity.
    /// </param>
    private static void InsertCellIntoSheet(SheetData sheetData, CellFormats cellFormats, ref uint rowIndex)
    {
        // Advance by 2 to leave a blank row between each style sample.
        rowIndex += 2;
        uint tempRowIndex = rowIndex;

        Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == tempRowIndex);
        if (row == null)
        {
            row = new Row() { RowIndex = rowIndex };
            sheetData.Append(row);
        }

        Cell cellA = new Cell()
        {
            CellReference = "A" + rowIndex,
            CellValue = new CellValue($"StyleIndex Id = {cellFormats.ChildElements.Count - 1}"),
            DataType = CellValues.String
        };
        row.Append(cellA);

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
