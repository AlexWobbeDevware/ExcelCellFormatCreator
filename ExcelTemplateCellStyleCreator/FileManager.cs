using System;
using System.Globalization;
using System.IO;

namespace ExcelTemplateCellStyleCreator
{
    public static class FileManager
    {
        public static void DeleteFileIfExists(string filePath, string culture)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                    Console.WriteLine(culture == "de" ? $"Vorhandene Datei '{filePath}' gelöscht." : $"Existing file '{filePath}' deleted.");
                }
            }
            catch (IOException ex)
            {
                Console.WriteLine(culture == "de" ? $"Fehler beim Löschen der Datei: {ex.Message}" : $"Error deleting file: {ex.Message}");
            }
        }
    }
}
