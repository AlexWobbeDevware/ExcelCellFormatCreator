using System;
using System.IO; // For file operations
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        string filePath = @"c:\temp\ExcelStyleTemplate.xlsx";

        // Delete the file if it already exists
        if (File.Exists(filePath))
        {
            File.Delete(filePath);
            Console.WriteLine($"Existing file '{filePath}' deleted.");
        }

        string lastFontName = "Calibri";
        double lastFontSize = 11;
        string lastFontColor = "000000";
        string lastBgColor = "FFFFFF";
        bool lastIsBold = false;
        bool lastIsItalic = false;

        using (SpreadsheetDocument document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
        {
            // Add a WorkbookPart to the document.
            WorkbookPart workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Styles" };
            sheets.Append(sheet);

            Worksheet worksheet = new Worksheet();

            // Gitterlinien ausblenden
            SheetViews sheetViews = new SheetViews();
            SheetView sheetView = new SheetView() { WorkbookViewId = (UInt32Value)0U, ShowGridLines = false };
            sheetViews.Append(sheetView);
            worksheet.Append(sheetViews);

            SheetData sheetData = new SheetData();
            worksheet.Append(sheetData);

            // Create and add styles
            WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            Stylesheet stylesheet = new Stylesheet();

            Fonts fonts = new Fonts(
                        new Font( // Default font
                            new FontSize() { Val = 11 },
                            new Color() { Theme = 1 },
                            new FontName() { Val = "Calibri" },
                            new FontFamilyNumbering() { Val = 2 },
                            new FontScheme() { Val = FontSchemeValues.Minor }
                        )
                    );

            Fills fills = new Fills(
                new Fill(new PatternFill() { PatternType = PatternValues.None }), // Default fill
                new Fill(new PatternFill(new ForegroundColor() { Rgb = new HexBinaryValue("FFFFFF") }) { PatternType = PatternValues.Solid }) // Background color
            );

            Borders borders = new Borders(
                new Border( // Default border
                    new LeftBorder(),
                    new RightBorder(),
                    new TopBorder(),
                    new BottomBorder(),
                    new DiagonalBorder()
                )
            );

            CellFormats cellFormats = new CellFormats(
                new CellFormat() // Default cell format
            );

            bool continueAdding = true;
            uint rowIndex = 2;
            uint columnIndex = 2;

            while (continueAdding)
            {
                Console.WriteLine("Neuen Stil hinzufügen:");

                // Beispielhafte Farbcodes anzeigen
                Console.WriteLine("Beispielhafte Farben: Rot: FF0000 | Grün: 00FF00 | Blau: 0000FF | Gelb: FFFF00 | Schwarz: 000000 | Weiß: FFFFFF");

                // Schriftart und Schriftgröße
                Console.WriteLine();
                Console.Write($"Schriftart (Standard: {lastFontName}): ");
                string fontName = Console.ReadLine() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(fontName))
                {
                    fontName = lastFontName;
                }
                lastFontName = fontName;

                Console.WriteLine();
                Console.Write($"Schriftgröße (Standard: {lastFontSize}): ");
                string fontSizeInput = Console.ReadLine() ?? string.Empty;
                double fontSize = string.IsNullOrWhiteSpace(fontSizeInput) ? lastFontSize : double.Parse(fontSizeInput);
                lastFontSize = fontSize;

                // Schriftfarbe
                Console.WriteLine();
                Console.Write($"Schriftfarbe (Standard: {lastFontColor} für Schwarz): ");
                string fontColor = Console.ReadLine() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(fontColor))
                {
                    fontColor = lastFontColor;
                }
                lastFontColor = fontColor;

                // Fett und Kursiv
                Console.WriteLine();
                Console.Write($"Fett (y/n, Standard: {(lastIsBold ? "y" : "n")}): ");
                string isBoldInput = Console.ReadLine() ?? string.Empty;
                bool isBold = string.IsNullOrWhiteSpace(isBoldInput) ? lastIsBold : isBoldInput.ToLower() == "y";
                lastIsBold = isBold;

                Console.WriteLine();
                Console.Write($"Kursiv (y/n, Standard: {(lastIsItalic ? "y" : "n")}): ");
                string isItalicInput = Console.ReadLine() ?? string.Empty;
                bool isItalic = string.IsNullOrWhiteSpace(isItalicInput) ? lastIsItalic : isItalicInput.ToLower() == "y";
                lastIsItalic = isItalic;

                // Hintergrundfarbe
                Console.WriteLine();
                Console.Write($"Hintergrundfarbe (Standard: {lastBgColor}): ");
                string bgColor = Console.ReadLine() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(bgColor))
                {
                    bgColor = lastBgColor;
                }
                lastBgColor = bgColor;

                // Rahmen (oulr für oben, unten, links, rechts; Standard: oulr): 
                Console.WriteLine();
                Console.Write("Rahmen (oulr für oben, unten, links, rechts; Standard: oulr, n für keine Rahmen): ");
                string borderInput = (Console.ReadLine() ?? string.Empty).ToLower();
                if (string.IsNullOrWhiteSpace(borderInput))
                {
                    borderInput = "oulr";
                }

                // Erstelle den Font
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
                fonts.Append(font);

                // Erstelle die Füllung
                Fill fill = new Fill(
                    new PatternFill(
                        new ForegroundColor() { Rgb = new HexBinaryValue(bgColor) }
                    )
                    { PatternType = PatternValues.Solid }
                );
                fills.Append(fill);

                // Erstelle den Rahmen
                Border border = new Border();

                if (borderInput != "n")
                {
                    if (borderInput.Contains("l"))
                        border.Append(new LeftBorder() { Style = BorderStyleValues.Thin });
                    if (borderInput.Contains("r"))
                        border.Append(new RightBorder() { Style = BorderStyleValues.Thin });
                    if (borderInput.Contains("o"))
                        border.Append(new TopBorder() { Style = BorderStyleValues.Thin });
                    if (borderInput.Contains("u"))
                        border.Append(new BottomBorder() { Style = BorderStyleValues.Thin });
                }

                // Check if the border already exists
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
                    borderId = (uint)borders.Count() - 1; // Update borderId to the new border's index
                }

                // Erstelle das Zellformat
                CellFormat cellFormat = new CellFormat()
                {
                    FontId = (UInt32)(fonts.Count() - 1),
                    FillId = (UInt32)(fills.Count() - 1),
                    BorderId = borderId,
                    ApplyFont = true,
                    ApplyFill = true,
                    ApplyBorder = true
                };
                cellFormats.Append(cellFormat);

                // Füge eine Zelle mit dem StyleIndex ein
                Row row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);

                Cell cell = new Cell()
                {
                    CellReference = $"{GetColumnName(columnIndex)}{rowIndex}",
                    CellValue = new CellValue("StyleIndex: " + (cellFormats.Count() - 1).ToString()),
                    DataType = CellValues.Number,
                    StyleIndex = (UInt32)(cellFormats.Count() - 1)
                };

                row.Append(cell);

                // Aktualisiere Zeilen- und Spaltenindex
                columnIndex += 2;
                if (columnIndex > 10)
                {
                    columnIndex = 1;
                    rowIndex += 2;
                }

                Console.WriteLine();
                Console.Write("Weiteren Stil hinzufügen? (y/n, Standard: y): ");
                string continueInput = Console.ReadLine() ?? string.Empty;
                continueAdding = string.IsNullOrWhiteSpace(continueInput) || continueInput.ToLower() == "y";
            }

            // Stylesheet zusammenstellen und speichern
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

            worksheetPart.Worksheet = worksheet;
            worksheetPart.Worksheet.Save();

            workbookPart.Workbook.Save();
        }

        Console.WriteLine("Excel-Datei wurde erfolgreich erstellt: " + filePath);
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
}