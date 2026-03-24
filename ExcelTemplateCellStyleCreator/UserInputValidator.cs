using DocumentFormat.OpenXml.Spreadsheet;
using ExcelTemplateCellStyleCreator.Core;
using System.Drawing.Text;
using static ExcelTemplateCellStyleCreator.Core.LocalizationHelper;

namespace ExcelTemplateCellStyleCreator
{
    /// <summary>
    /// Validates and normalizes all interactive user input collected during style creation.
    /// Each method loops until the user provides an acceptable value.
    /// </summary>
    public static class UserInputValidator
    {
        // Cached once at startup to avoid enumerating installed fonts on every validation call.
        private static readonly InstalledFontCollection InstalledFonts = new InstalledFontCollection();

        /// <summary>
        /// Ensures <paramref name="fontName"/> matches an installed font family (case-insensitive).
        /// Prompts the user to re-enter until a valid name is provided.
        /// </summary>
        /// <param name="fontName">Initial font name to validate.</param>
        /// <param name="culture">Two-letter ISO language name for localized error messages.</param>
        /// <returns>A valid, installed font family name.</returns>
        public static string ValidateFontName(string fontName, string culture)
        {
            var fontFamilies = InstalledFonts.Families;

            while (!fontFamilies.Any(f => f.Name.Equals(fontName, StringComparison.OrdinalIgnoreCase)))
            {
                Console.Write(Localize(culture,
                    "Ungültiger Schriftartname. Bitte geben Sie einen gültigen Schriftartnamen ein: ",
                    "Invalid font name. Please enter a valid font name: "));
                fontName = Console.ReadLine();
            }

            return fontName;
        }

        /// <summary>
        /// Ensures <paramref name="fontSizeInput"/> can be parsed as a positive number.
        /// Prompts the user to re-enter until a valid value is provided.
        /// </summary>
        /// <param name="fontSizeInput">Initial font size string to validate.</param>
        /// <param name="culture">Two-letter ISO language name for localized error messages.</param>
        /// <returns>A string representation of a valid positive font size.</returns>
        public static string ValidateFontSize(string fontSizeInput, string culture)
        {
            while (!double.TryParse(fontSizeInput, out double fontSize) || fontSize <= 0)
            {
                Console.Write(Localize(culture,
                    "Ungültige Schriftgröße. Bitte geben Sie eine positive Zahl ein: ",
                    "Invalid font size. Please enter a positive number: "));
                fontSizeInput = Console.ReadLine();
            }

            return fontSizeInput;
        }

        /// <summary>
        /// Ensures <paramref name="colorInput"/> is exactly 6 hexadecimal characters (e.g. "FF0000").
        /// Prompts the user to re-enter until a valid value is provided.
        /// </summary>
        /// <param name="colorInput">Initial hex color string to validate.</param>
        /// <param name="culture">Two-letter ISO language name for localized error messages.</param>
        /// <returns>A valid 6-digit hex color string.</returns>
        public static string ValidateHexColor(string colorInput, string culture)
        {
            while (!HexColorValidator.IsValid(colorInput))
            {
                Console.Write(Localize(culture,
                    "Ungültige Farbe. Bitte geben Sie einen gültigen 6-stelligen Hex-Farbcode ein: ",
                    "Invalid color. Please enter a valid 6-digit hex color code: "));
                colorInput = Console.ReadLine();
            }

            return colorInput;
        }

        /// <summary>
        /// Converts a yes/no string to a boolean. Accepts "y", "j" (German) as true and "n", "nein" as false.
        /// Prompts the user to re-enter until a recognized value is provided.
        /// </summary>
        /// <param name="input">Initial input string to validate.</param>
        /// <param name="culture">Two-letter ISO language name for localized error messages.</param>
        /// <returns><see langword="true"/> for "y" or "j"; <see langword="false"/> for "n" or "nein".</returns>
        public static bool ValidateYesNoInput(string input, string culture)
        {
            input = input.Trim().ToLower();
            while (input != "y" && input != "n" && input != "j" && input != "nein")
            {
                Console.Write(Localize(culture,
                    "Ungültige Eingabe. Bitte geben Sie 'y' oder 'n' ein: ",
                    "Invalid input. Please enter 'y' or 'n': "));
                input = Console.ReadLine()?.Trim().ToLower();
            }

            return input == "y" || input == "j";
        }

        /// <summary>
        /// Ensures <paramref name="borderInput"/> is a non-empty string containing only
        /// the characters 'l', 'r', 't', 'b' (left, right, top, bottom).
        /// Prompts the user to re-enter until a valid value is provided.
        /// </summary>
        /// <param name="borderInput">Initial border selection string to validate.</param>
        /// <param name="culture">Two-letter ISO language name for localized error messages.</param>
        /// <returns>A valid border selection string.</returns>
        public static string ValidateBorderSelection(string borderInput, string culture)
        {
            while (string.IsNullOrWhiteSpace(borderInput) || !borderInput.All(c => "lrtb".Contains(c)))
            {
                Console.Write(Localize(culture,
                    "Ungültige Rahmeneingabe. Bitte geben Sie eine gültige Kombination aus 'l', 'r', 't', 'b' ein: ",
                    "Invalid border selection. Please enter a valid combination of 'l', 'r', 't', 'b': "));
                borderInput = Console.ReadLine();
            }

            return borderInput;
        }

        /// <summary>
        /// Converts a single-character alignment code to the corresponding
        /// <see cref="HorizontalAlignmentValues"/> enum value.
        /// Accepts 'l' (Left), 'c' (Center), 'r' (Right).
        /// Prompts the user to re-enter until a recognized code is provided.
        /// </summary>
        /// <param name="alignmentInput">Initial alignment code to interpret.</param>
        /// <param name="culture">Two-letter ISO language name for localized error messages.</param>
        /// <returns>The matching <see cref="HorizontalAlignmentValues"/>.</returns>
        public static HorizontalAlignmentValues GetHorizontalAlignment(string alignmentInput, string culture)
        {
            alignmentInput = alignmentInput.ToLower();

            while (true)
            {
                switch (alignmentInput)
                {
                    case "l": return HorizontalAlignmentValues.Left;
                    case "c": return HorizontalAlignmentValues.Center;
                    case "r": return HorizontalAlignmentValues.Right;
                }

                Console.Write(Localize(culture,
                    "Ungültige Eingabe. Bitte geben Sie 'l' für Links, 'c' für Zentrum oder 'r' für Rechts ein: ",
                    "Invalid input. Please enter 'l' for Left, 'c' for Center, or 'r' for Right: "));
                alignmentInput = Console.ReadLine()?.ToLower();
            }
        }

        /// <summary>
        /// Converts a single-character alignment code to the corresponding
        /// <see cref="VerticalAlignmentValues"/> enum value.
        /// Accepts 't' (Top), 'c' (Center), 'b' (Bottom).
        /// Prompts the user to re-enter until a recognized code is provided.
        /// </summary>
        /// <param name="alignmentInput">Initial alignment code to interpret.</param>
        /// <param name="culture">Two-letter ISO language name for localized error messages.</param>
        /// <returns>The matching <see cref="VerticalAlignmentValues"/>.</returns>
        public static VerticalAlignmentValues GetVerticalAlignment(string alignmentInput, string culture)
        {
            alignmentInput = alignmentInput.ToLower();

            while (true)
            {
                switch (alignmentInput)
                {
                    case "t": return VerticalAlignmentValues.Top;
                    case "c": return VerticalAlignmentValues.Center;
                    case "b": return VerticalAlignmentValues.Bottom;
                }

                Console.Write(Localize(culture,
                    "Ungültige Eingabe. Bitte geben Sie 't' für Oben, 'c' für Mitte oder 'b' für Unten ein: ",
                    "Invalid input. Please enter 't' for Top, 'c' for Center, or 'b' for Bottom: "));
                alignmentInput = Console.ReadLine()?.ToLower();
            }
        }
    }
}
