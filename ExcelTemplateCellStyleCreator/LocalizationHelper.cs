namespace ExcelTemplateCellStyleCreator
{
    /// <summary>
    /// Provides simple culture-based string selection for German and English UI messages.
    /// Uses the two-letter ISO language name from <see cref="System.Globalization.CultureInfo"/>
    /// to decide which string to return.
    /// </summary>
    public static class LocalizationHelper
    {
        /// <summary>
        /// Returns the German or English string based on the given culture code.
        /// </summary>
        /// <param name="culture">Two-letter ISO language name (e.g. "de" or "en").</param>
        /// <param name="de">String to return when <paramref name="culture"/> is "de".</param>
        /// <param name="en">String to return for any other culture.</param>
        /// <returns>
        /// <paramref name="de"/> if <paramref name="culture"/> equals "de";
        /// otherwise <paramref name="en"/>.
        /// </returns>
        public static string Localize(string culture, string de, string en)
        {
            return culture == "de" ? de : en;
        }
    }
}
