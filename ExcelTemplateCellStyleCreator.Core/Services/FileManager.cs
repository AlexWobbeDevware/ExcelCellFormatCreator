namespace ExcelTemplateCellStyleCreator.Core
{
    /// <summary>
    /// Handles file system operations for the output Excel file.
    /// </summary>
    public static class FileManager
    {
        /// <summary>
        /// Deletes the file at <paramref name="filePath"/> if it exists.
        /// Returns true if the file was deleted or did not exist; false if deletion failed.
        /// </summary>
        public static bool DeleteFileIfExists(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                    File.Delete(filePath);
                return true;
            }
            catch (IOException)
            {
                return false;
            }
        }
    }
}
