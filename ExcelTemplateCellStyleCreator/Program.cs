using System;
using System.Globalization;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

class Program
{

    static void Main(string[] args)
    {
        string filePath = @"c:\temp\ExcelStyleTemplate.xlsx";
        var culture = CultureInfo.CurrentCulture.TwoLetterISOLanguageName;

        DeleteFileIfExists(filePath, culture);

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

            Fonts fonts = CreateDefaultFonts();
            Fills fills = CreateDefaultFills();
            Borders borders = CreateDefaultBorders();
            CellFormats cellFormats = new CellFormats(new CellFormat());

            bool continueAdding = true;
            uint rowIndex = 2;
            uint columnIndex = 2;

            while (continueAdding)
            {
                Console.WriteLine(culture == "de" ? "Neuen Stil hinzufügen:" : "Add a new style:");
                Console.WriteLine(culture == "de" ? "Beispielhafte Farben: Rot: FF0000 | Grün: 00FF00 | Blau: 0000FF | Gelb: FFFF00 | Schwarz: 000000 | Weiß: FFFFFF"
                                                  : "Example colors: Red: FF0000 | Green: 00FF00 | Blue: 0000FF | Yellow: FFFF00 | Black: 000000 | White: FFFFFF");

                lastFontName = GetUserInput(culture, "Schriftart", "Font name", lastFontName);
                lastFontSize = double.Parse(GetUserInput(culture, "Schriftgröße", "Font size", lastFontSize.ToString()));
                lastFontColor = GetUserInput(culture, "Schriftfarbe", "Font color", lastFontColor);
                lastIsBold = GetUserInput(culture, "Fett (y/n)", "Bold (y/n)", lastIsBold ? "y" : "n").ToLower() == "y";
                lastIsItalic = GetUserInput(culture, "Kursiv (y/n)", "Italic (y/n)", lastIsItalic ? "y" : "n").ToLower() == "y";
                lastBgColor = GetUserInput(culture, "Hintergrundfarbe", "Background color", lastBgColor);
                lastBorderSelection = GetUserInput(culture, "Rahmen auswählen (left, right, top, bottom)", "Select borders (left, right, top, bottom)", lastBorderSelection);

                ConfigureFont(lastFontName, lastFontSize, lastFontColor, lastIsBold, lastIsItalic, fonts);
                ConfigureFills(lastBgColor, fills);
                Border border = ConfigureBorder(culture, lastBorderSelection);
                bool configureAlignment = GetUserInput(culture, "Textausrichtung und Umbruch konfigurieren", "Configure text alignment and wrapping", "n").ToLower() == "y";
                if (configureAlignment)
                {
                    lastHorizontalAlignment = GetHorizontalAlignment(culture, lastHorizontalAlignment);
                    lastVerticalAlignment = GetVerticalAlignment(culture, lastVerticalAlignment);
                    lastWrapText = GetUserInput(culture, "Textumbruch aktivieren", "Enable text wrapping", lastWrapText ? "y" : "n").ToLower() == "y";
                }
                uint borderId = GetOrCreateBorderId(borders, border);

                CellFormat cellFormat = CreateCellFormat(fonts, fills, borderId, configureAlignment, lastHorizontalAlignment, lastVerticalAlignment, lastWrapText);
                cellFormats.Append(cellFormat);

                InsertCellIntoSheet(sheetData, cellFormats, ref rowIndex, ref columnIndex);

                continueAdding = GetUserInput(culture, "Weiteren Stil hinzufügen", "Add another style (y/n)", "y").ToLower() == "y";
            }

            SaveStylesheet(stylesPart, stylesheet, fonts, fills, borders, cellFormats);
            worksheetPart.Worksheet = worksheet;
            worksheetPart.Worksheet.Save();
            workbookPart.Workbook.Save();
        }

        Console.WriteLine(culture == "de" ? $"Excel-Datei wurde erfolgreich erstellt: {filePath}" : $"Excel file successfully created: {filePath}");
    }

    private static void ConfigureFills(string lastBgColor, Fills fills)
    {

        // Erstelle die Füllung
        Fill fill = new Fill(
            new PatternFill(
                new ForegroundColor() { Rgb = new HexBinaryValue(lastBgColor) }
            )
            { PatternType = PatternValues.Solid }
        );
        fills.Append(fill);
    }

    private static void ConfigureFont(string lastFontName, double lastFontSize, string lastFontColor, bool lastIsBold, bool lastIsItalic, Fonts fonts)
    {
        // Erstelle den Font
        Font font = new Font();
        font.Append(new FontSize() { Val = lastFontSize });
        font.Append(new Color() { Rgb = new HexBinaryValue() { Value = lastFontColor } });
        font.Append(new FontName() { Val = lastFontName });
        if (lastIsBold)
        {
            font.Append(new Bold());
        }
        if (lastIsItalic)
        {
            font.Append(new Italic());
        }
        fonts.Append(font);
    }

    private static void DeleteFileIfExists(string filePath, string culture)
    {
        if (File.Exists(filePath))
        {
            File.Delete(filePath);
            Console.WriteLine(culture == "de" ? $"Vorhandene Datei '{filePath}' gelöscht." : $"Existing file '{filePath}' deleted.");
        }
    }

    private static void HideGridLines(Worksheet worksheet)
    {
        SheetViews sheetViews = new SheetViews();
        SheetView sheetView = new SheetView() { WorkbookViewId = (UInt32Value)0U, ShowGridLines = false };
        sheetViews.Append(sheetView);
        worksheet.Append(sheetViews);
    }

    private static Fonts CreateDefaultFonts()
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

    private static Fills CreateDefaultFills()
    {
        return new Fills(
            new Fill(new PatternFill() { PatternType = PatternValues.None }),
            new Fill(new PatternFill(new ForegroundColor() { Rgb = new HexBinaryValue("FFFFFF") }) { PatternType = PatternValues.Solid })
        );
    }

    private static Borders CreateDefaultBorders()
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

    private static string GetUserInput(string culture, string promptDe, string promptEn, string defaultValue)
    {
        Console.WriteLine();
        Console.Write(culture == "de" ? $"{promptDe} (Standard: {defaultValue}): " : $"{promptEn} (Default: {defaultValue}): ");
        string input = Console.ReadLine() ?? string.Empty;
        return string.IsNullOrWhiteSpace(input) ? defaultValue : input;
    }

    private static Border ConfigureBorder(string culture, string borderSelection)
    {
        Border border = new Border();

        if (borderSelection.Contains("l"))
            border.Append(new LeftBorder(new Color() { Rgb = new HexBinaryValue("000000") }) { Style = BorderStyleValues.Thin });
        if (borderSelection.Contains("r"))
            border.Append(new RightBorder(new Color() { Rgb = new HexBinaryValue("000000") }) { Style = BorderStyleValues.Thin });
        if (borderSelection.Contains("t"))
            border.Append(new TopBorder(new Color() { Rgb = new HexBinaryValue("000000") }) { Style = BorderStyleValues.Thin });
        if (borderSelection.Contains("b"))
            border.Append(new BottomBorder(new Color() { Rgb = new HexBinaryValue("000000") }) { Style = BorderStyleValues.Thin });

        if (!string.IsNullOrWhiteSpace(borderSelection))
        {
            bool configureBorders = GetUserInput(culture, "Rahmen im Detail konfigurieren", "Configure borders in detail", "n").ToLower() == "y";
            if (configureBorders)
            {
                ConfigureDetailedBorders(culture, border, borderSelection);
            }
        }

        return border;
    }

    private static void ConfigureDetailedBorders(string culture, Border border, string borderSelection)
    {
        if (borderSelection.Contains("l"))
            border.LeftBorder = ConfigureLeftBorder(culture);
        if (borderSelection.Contains("r"))
            border.RightBorder = ConfigureRightBorder(culture);
        if (borderSelection.Contains("t"))
            border.TopBorder = ConfigureTopBorder(culture);
        if (borderSelection.Contains("b"))
            border.BottomBorder = ConfigureBottomBorder(culture);
    }

    private static LeftBorder ConfigureLeftBorder(string culture)
    {
        Console.Write(culture == "de" ? $"Links Rahmenstil wählen (thin = 1/medium = 2/thick = 3/dashed = 4/dotted = 5, Standard: thin): "
                                      : $"Left border style (thin = 1/medium = 2/thick = 3/dashed = 4/dotted = 5, Default: thin): ");
        string borderStyleInput = Console.ReadLine()?.ToLower() ?? "1";
        BorderStyleValues borderStyle = borderStyleInput switch
        {
            "2" => BorderStyleValues.Medium,
            "3" => BorderStyleValues.Thick,
            "4" => BorderStyleValues.Dashed,
            "5" => BorderStyleValues.Dotted,
            _ => BorderStyleValues.Thin,
        };

        Console.Write(culture == "de" ? $"Links Rahmenfarbe (Hex-Wert, Standard: 000000 für Schwarz): "
                                      : $"Left border color (Hex value, Default: 000000 for Black): ");
        string borderColor = Console.ReadLine() ?? "000000";

        return new LeftBorder(new Color() { Rgb = new HexBinaryValue(borderColor) }) { Style = borderStyle };
    }

    private static RightBorder ConfigureRightBorder(string culture)
    {
        Console.Write(culture == "de" ? $"Rechts Rahmenstil wählen (thin = 1/medium = 2/thick = 3/dashed = 4/dotted = 5, Standard: thin): "
                                      : $"Right border style (thin = 1/medium = 2/thick = 3/dashed = 4/dotted = 5, Default: thin): ");
        string borderStyleInput = Console.ReadLine()?.ToLower() ?? "1";
        BorderStyleValues borderStyle = borderStyleInput switch
        {
            "2" => BorderStyleValues.Medium,
            "3" => BorderStyleValues.Thick,
            "4" => BorderStyleValues.Dashed,
            "5" => BorderStyleValues.Dotted,
            _ => BorderStyleValues.Thin,
        };

        Console.Write(culture == "de" ? $"Rechts Rahmenfarbe (Hex-Wert, Standard: 000000 für Schwarz): "
                                      : $"Right border color (Hex value, Default: 000000 for Black): ");
        string borderColor = Console.ReadLine() ?? "000000";

        return new RightBorder(new Color() { Rgb = new HexBinaryValue(borderColor) }) { Style = borderStyle };
    }

    private static TopBorder ConfigureTopBorder(string culture)
    {
        Console.Write(culture == "de" ? $"Oben Rahmenstil wählen (thin = 1/medium = 2/thick = 3/dashed = 4/dotted = 5, Standard: thin): "
                                      : $"Top border style (thin = 1/medium = 2/thick = 3/dashed = 4/dotted = 5, Default: thin): ");
        string borderStyleInput = Console.ReadLine()?.ToLower() ?? "1";
        BorderStyleValues borderStyle = borderStyleInput switch
        {
            "2" => BorderStyleValues.Medium,
            "3" => BorderStyleValues.Thick,
            "4" => BorderStyleValues.Dashed,
            "5" => BorderStyleValues.Dotted,
            _ => BorderStyleValues.Thin,
        };

        Console.Write(culture == "de" ? $"Oben Rahmenfarbe (Hex-Wert, Standard: 000000 für Schwarz): "
                                      : $"Top border color (Hex value, Default: 000000 for Black): ");
        string borderColor = Console.ReadLine() ?? "000000";

        return new TopBorder(new Color() { Rgb = new HexBinaryValue(borderColor) }) { Style = borderStyle };
    }

    private static BottomBorder ConfigureBottomBorder(string culture)
    {
        Console.Write(culture == "de" ? $"Unten Rahmenstil wählen (thin = 1/medium = 2/thick = 3/dashed = 4/dotted = 5, Standard: thin): "
                                      : $"Bottom border style (thin = 1/medium = 2/thick = 3/dashed = 4/dotted = 5, Default: thin): ");
        string borderStyleInput = Console.ReadLine()?.ToLower() ?? "1";
        BorderStyleValues borderStyle = borderStyleInput switch
        {
            "2" => BorderStyleValues.Medium,
            "3" => BorderStyleValues.Thick,
            "4" => BorderStyleValues.Dashed,
            "5" => BorderStyleValues.Dotted,
            _ => BorderStyleValues.Thin,
        };

        Console.Write(culture == "de" ? $"Unten Rahmenfarbe (Hex-Wert, Standard: 000000 für Schwarz): "
                                      : $"Bottom border color (Hex value, Default: 000000 for Black): ");
        string borderColor = Console.ReadLine() ?? "000000";

        return new BottomBorder(new Color() { Rgb = new HexBinaryValue(borderColor) }) { Style = borderStyle };
    }

    private static HorizontalAlignmentValues GetHorizontalAlignment(string culture, HorizontalAlignmentValues defaultAlignment)
    {
        Console.WriteLine();
        Console.Write(culture == "de" ? $"Horizontale Ausrichtung (l=links, c=zentrisch, r=rechts, Standard: {defaultAlignment.ToString().ToLower()[0]}): "
                                      : $"Horizontal alignment (l=left, c=center, r=right, Default: {defaultAlignment.ToString().ToLower()[0]}): ");
        string horizontalAlignmentInput = Console.ReadLine() ?? defaultAlignment.ToString().ToLower()[0].ToString();
        return horizontalAlignmentInput switch
        {
            "c" => HorizontalAlignmentValues.Center,
            "r" => HorizontalAlignmentValues.Right,
            _ => HorizontalAlignmentValues.Left,
        };
    }

    private static VerticalAlignmentValues GetVerticalAlignment(string culture, VerticalAlignmentValues defaultAlignment)
    {
        Console.WriteLine();
        Console.Write(culture == "de" ? $"Vertikale Ausrichtung (t=oben, m=mittig, b=unten, Standard: {defaultAlignment.ToString().ToLower()[0]}): "
                                      : $"Vertical alignment (t=top, m=middle, b=bottom, Default: {defaultAlignment.ToString().ToLower()[0]}): ");
        string verticalAlignmentInput = Console.ReadLine() ?? defaultAlignment.ToString().ToLower()[0].ToString();
        return verticalAlignmentInput switch
        {
            "t" => VerticalAlignmentValues.Top,
            "b" => VerticalAlignmentValues.Bottom,
            _ => VerticalAlignmentValues.Center,
        };
    }

    private static uint GetOrCreateBorderId(Borders borders, Border border)
    {
        uint borderId = 0;
        bool borderExists = false;
        foreach (Border existingBorder in borders.Elements<Border>())
        {
            if (AreBordersEqual(existingBorder, border))
            {
                borderExists = true;
                break;
            }
            borderId++;
        }

        if (!borderExists)
        {
            borders.Append(border);
            borderId = (uint)borders.Count() - 1;
        }

        return borderId;
    }

    private static CellFormat CreateCellFormat(Fonts fonts, Fills fills, uint borderId, bool configureAlignment, HorizontalAlignmentValues horizontalAlignment, VerticalAlignmentValues verticalAlignment, bool wrapText)
    {
        CellFormat cellFormat = new CellFormat()
        {
            FontId = (UInt32)(fonts.Count() - 1),
            FillId = (UInt32)(fills.Count() - 1),
            BorderId = borderId,
            ApplyFont = true,
            ApplyFill = true,
            ApplyBorder = true,
            ApplyAlignment = configureAlignment
        };

        if (configureAlignment)
        {
            cellFormat.Alignment = new Alignment()
            {
                Horizontal = horizontalAlignment,
                Vertical = verticalAlignment,
                WrapText = wrapText
            };
        }

        return cellFormat;
    }

    private static void InsertCellIntoSheet(SheetData sheetData, CellFormats cellFormats, ref uint rowIndex, ref uint columnIndex)
    {
        Row row = new Row() { RowIndex = rowIndex };
        sheetData.Append(row);

        Cell cell = new Cell()
        {
            CellReference = $"{GetColumnName(columnIndex)}{rowIndex}",
            CellValue = new CellValue((cellFormats.Count() - 1).ToString()),
            DataType = CellValues.Number,
            StyleIndex = (UInt32)(cellFormats.Count() - 1)
        };

        row.Append(cell);

        columnIndex += 2;
        if (columnIndex > 10)
        {
            columnIndex = 1;
            rowIndex += 2;
        }
    }

    static string GetColumnName(uint index)
    {
        const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        string columnName = "";

        while (index > 0)
        {
            uint remainder = (index - 1) % 26;
            columnName = letters[(int)remainder] + columnName;
            index = (index - remainder) / 26;
        }

        return columnName;
    }

    private static bool AreBordersEqual(Border border1, Border border2)
    {
        if (border1 == null || border2 == null)
            return false;

        return border1.LeftBorder?.Style == border2.LeftBorder?.Style &&
               border1.RightBorder?.Style == border2.RightBorder?.Style &&
               border1.TopBorder?.Style == border2.TopBorder?.Style &&
               border1.BottomBorder?.Style == border2.BottomBorder?.Style &&
               border1.DiagonalBorder?.Style == border2.DiagonalBorder?.Style;
    }

    private static void SaveStylesheet(WorkbookStylesPart stylesPart, Stylesheet stylesheet, Fonts fonts, Fills fills, Borders borders, CellFormats cellFormats)
    {
        stylesheet.Fonts = new Fonts();
        foreach (var font in fonts.Elements<Font>())
        {
            stylesheet.Fonts.Append(font.CloneNode(true));
        }

        stylesheet.Fills = new Fills();
        foreach (var fill in fills.Elements<Fill>())
        {
            stylesheet.Fills.Append(fill.CloneNode(true));
        }

        stylesheet.Borders = new Borders();
        foreach (var border in borders.Elements<Border>())
        {
            stylesheet.Borders.Append(border.CloneNode(true));
        }

        stylesheet.CellFormats = new CellFormats();
        foreach (var cellFormat in cellFormats.Elements<CellFormat>())
        {
            stylesheet.CellFormats.Append(cellFormat.CloneNode(true));
        }

        stylesPart.Stylesheet = stylesheet;
        stylesPart.Stylesheet.Save();
    }
}