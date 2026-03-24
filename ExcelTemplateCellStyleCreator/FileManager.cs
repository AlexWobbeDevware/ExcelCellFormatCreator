using static ExcelTemplateCellStyleCreator.LocalizationHelper;

namespace ExcelTemplateCellStyleCreator
{
    /// <summary>
    /// Handles file system operations for the output Excel file.
    /// </summary>
    public static class FileManager
    {
        /// <summary>
        /// Deletes the file at <paramref name="filePath"/> if it exists, so that a fresh
        /// workbook can be written without conflicts.
        /// Does nothing if the file is absent. Logs a confirmation or error message to the console.
        /// </summary>
        /// <param name="filePath">Absolute path of the file to delete.</param>
        /// <param name="culture">Two-letter ISO language name used to localize console output.</param>
        public static void DeleteFileIfExists(string filePath, string culture)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                    Console.WriteLine(Localize(culture,
                        $"Vorhandene Datei '{filePath}' gelöscht.",
                        $"Existing file '{filePath}' deleted."));
                }
            }
            catch (IOException ex)
            {
                Console.WriteLine(Localize(culture,
                    $"Fehler beim Löschen der Datei: {ex.Message}",
                    $"Error deleting file: {ex.Message}"));
            }
        }
    }
}
