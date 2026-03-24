using System.Text.RegularExpressions;

namespace ExcelTemplateCellStyleCreator.Core
{
    /// <summary>
    /// Validates and normalizes 6-digit hex color strings.
    /// </summary>
    public static class HexColorValidator
    {
        /// <summary>
        /// Returns true if <paramref name="hex"/> is a valid 6-digit hex color (e.g. "FF0000").
        /// </summary>
        public static bool IsValid(string hex)
            => !string.IsNullOrEmpty(hex) && Regex.IsMatch(hex, StyleConstants.HexColorPattern);

        /// <summary>
        /// Normalizes a hex color string: trims whitespace, strips 8-char ARGB alpha prefix,
        /// and converts to uppercase. Returns "000000" if input is null or empty.
        /// </summary>
        public static string Normalize(string? hex)
        {
            if (string.IsNullOrWhiteSpace(hex)) return StyleConstants.DefaultFontColor;
            hex = hex.Trim().TrimStart('#');
            if (hex.Length == 8) hex = hex[2..];
            return hex.ToUpper();
        }
    }
}
