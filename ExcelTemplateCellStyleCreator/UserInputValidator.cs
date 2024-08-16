using DocumentFormat.OpenXml.Spreadsheet;
using System.Drawing.Text;

namespace ExcelTemplateCellStyleCreator
{
    public static class UserInputValidator
    {
        public static string ValidateFontName(string fontName, string culture)
        {
            InstalledFontCollection installedFonts = new InstalledFontCollection();
            System.Drawing.FontFamily[] fontFamilies = installedFonts.Families;

            while (!fontFamilies.Any(f => f.Name.Equals(fontName, StringComparison.OrdinalIgnoreCase)))
            {
                Console.Write(culture == "de"
                    ? "Ungültiger Schriftartname. Bitte geben Sie einen gültigen Schriftartnamen ein: "
                    : "Invalid font name. Please enter a valid font name: ");
                fontName = Console.ReadLine();
            }

            return fontName;
        }

        public static string ValidateFontSize(string fontSizeInput, string culture)
        {
            while (!double.TryParse(fontSizeInput, out double fontSize) || fontSize <= 0)
            {
                Console.Write(culture == "de"
                    ? "Ungültige Schriftgröße. Bitte geben Sie eine positive Zahl ein: "
                    : "Invalid font size. Please enter a positive number: ");
                fontSizeInput = Console.ReadLine();
            }

            return fontSizeInput;
        }

        public static string ValidateHexColor(string colorInput, string culture)
        {
            while (!System.Text.RegularExpressions.Regex.IsMatch(colorInput, @"^[0-9A-Fa-f]{6}$"))
            {
                Console.Write(culture == "de"
                    ? "Ungültige Farbe. Bitte geben Sie einen gültigen 6-stelligen Hex-Farbcode ein: "
                    : "Invalid color. Please enter a valid 6-digit hex color code: ");
                colorInput = Console.ReadLine();
            }

            return colorInput;
        }

        public static bool ValidateYesNoInput(string input, string culture)
        {
            input = input.Trim().ToLower();
            while (input != "y" && input != "n" && input != "j" && input != "nein")
            {
                Console.Write(culture == "de"
                    ? "Ungültige Eingabe. Bitte geben Sie 'y' oder 'n' ein: "
                    : "Invalid input. Please enter 'y' or 'n': ");
                input = Console.ReadLine()?.Trim().ToLower();
            }

            return input == "y" || input == "j";
        }

        public static string ValidateBorderSelection(string borderInput, string culture)
        {
            while (string.IsNullOrWhiteSpace(borderInput) || !borderInput.All(c => "lrtb".Contains(c)))
            {
                Console.Write(culture == "de"
                    ? "Ungültige Rahmeneingabe. Bitte geben Sie eine gültige Kombination aus 'l', 'r', 't', 'b' ein: "
                    : "Invalid border selection. Please enter a valid combination of 'l', 'r', 't', 'b': ");
                borderInput = Console.ReadLine();
            }

            return borderInput;
        }

        public static HorizontalAlignmentValues GetHorizontalAlignment(string alignmentInput, string culture)
        {
            alignmentInput = alignmentInput.ToLower();

            switch (alignmentInput)
            {
                case "l":
                    return HorizontalAlignmentValues.Left;
                case "c":
                    return HorizontalAlignmentValues.Center;
                case "r":
                    return HorizontalAlignmentValues.Right;
                default:
                    Console.Write(culture == "de"
                        ? "Ungültige Eingabe. Bitte geben Sie 'l' für Links, 'c' für Zentrum oder 'r' für Rechts ein: "
                        : "Invalid input. Please enter 'l' for Left, 'c' for Center, or 'r' for Right: ");
                    alignmentInput = Console.ReadLine()?.ToLower();
                    return GetHorizontalAlignment(alignmentInput, culture); // Recursively call to handle incorrect input
            }
        }

        public static VerticalAlignmentValues GetVerticalAlignment(string alignmentInput, string culture)
        {
            alignmentInput = alignmentInput.ToLower();

            switch (alignmentInput)
            {
                case "t":
                    return VerticalAlignmentValues.Top;
                case "c":
                    return VerticalAlignmentValues.Center;
                case "b":
                    return VerticalAlignmentValues.Bottom;
                default:
                    Console.Write(culture == "de"
                        ? "Ungültige Eingabe. Bitte geben Sie 't' für Oben, 'c' für Mitte oder 'b' für Unten ein: "
                        : "Invalid input. Please enter 't' for Top, 'c' for Center, or 'b' for Bottom: ");
                    alignmentInput = Console.ReadLine()?.ToLower();
                    return GetVerticalAlignment(alignmentInput, culture); // Recursively call to handle incorrect input
            }
        }
    }
}
