using System.Drawing.Text;

namespace ExcelTemplateCellStyleCreator
{
    public static class UserInputValidator
    {
        public static string ValidateFontName(string fontName)
        {
            InstalledFontCollection installedFonts = new InstalledFontCollection();
            System.Drawing.FontFamily[] fontFamilies = installedFonts.Families;

            while (!fontFamilies.Any(f => f.Name.Equals(fontName, StringComparison.OrdinalIgnoreCase)))
            {
                Console.WriteLine("Invalid font name. Please enter a valid font name:");
                fontName = Console.ReadLine();
            }

            return fontName;
        }

        public static string ValidateFontSize(string fontSizeInput)
        {
            while (!double.TryParse(fontSizeInput, out double fontSize) || fontSize <= 0)
            {
                Console.WriteLine("Invalid font size. Please enter a positive number:");
                fontSizeInput = Console.ReadLine();
            }

            return fontSizeInput;
        }

        public static string ValidateHexColor(string colorInput)
        {
            while (!System.Text.RegularExpressions.Regex.IsMatch(colorInput, @"^[0-9A-Fa-f]{6}$"))
            {
                Console.WriteLine("Invalid color. Please enter a valid 6-digit hex color code:");
                colorInput = Console.ReadLine();
            }

            return colorInput;
        }

        public static bool ValidateYesNoInput(string input, string culture)
        {
            input = input.Trim().ToLower();
            while (input != "y" && input != "n" && input != "j" && input != "nein")
            {
                Console.WriteLine(culture == "de" ? "Ungültige Eingabe. Bitte geben Sie 'y' oder 'n' ein:" : "Invalid input. Please enter 'y' or 'n':");
                input = Console.ReadLine()?.Trim().ToLower();
            }

            return input == "y" || input == "j";
        }

        public static string ValidateBorderSelection(string borderInput)
        {
            while (string.IsNullOrWhiteSpace(borderInput) || !borderInput.All(c => "lrtb".Contains(c)))
            {
                Console.WriteLine("Invalid border selection. Please enter a valid combination of 'l', 'r', 't', 'b':");
                borderInput = Console.ReadLine();
            }

            return borderInput;
        }
    }

}
