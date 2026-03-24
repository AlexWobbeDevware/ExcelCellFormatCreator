namespace ExcelTemplateCellStyleCreator.Core
{
    /// <summary>
    /// Centralizes all magic values used across the application.
    /// </summary>
    public static class StyleConstants
    {
        public const string HexColorPattern = @"^[0-9A-Fa-f]{6}$";

        public const string DefaultFontName = "Calibri";
        public const double DefaultFontSize = 11.0;
        public const string DefaultFontColor = "000000";
        public const string DefaultBgColor = "FFFFFF";
        public const string DefaultBorderSelection = "lrtb";

        public const string DefaultFilePath = @"C:\temp\ExcelStyleTemplate.xlsx";
        public const string SheetName = "Styles";
        public const uint RowSpacing = 2;
    }
}
