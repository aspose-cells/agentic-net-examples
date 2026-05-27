using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookEncryptionDemo
{
    public static class WorkbookEncryptionHelper
    {
        /// <summary>
        /// Determines whether the specified workbook file is encrypted (requires a password to open).
        /// </summary>
        /// <param name="filePath">Full path to the workbook file.</param>
        /// <returns>True if the workbook is encrypted; otherwise, false.</returns>
        public static bool IsWorkbookEncrypted(string filePath)
        {
            // Ensure the file exists before attempting detection
            if (!File.Exists(filePath))
                throw new FileNotFoundException("Workbook file not found.", filePath);

            // Detect the file format and retrieve encryption information without loading the workbook.
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
            return formatInfo.IsEncrypted;
        }
    }

    internal class Program
    {
        static void Main(string[] args)
        {
            // Example usage: pass workbook path as first argument
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the full path to the workbook file as an argument.");
                return;
            }

            string workbookPath = args[0];

            try
            {
                bool encrypted = WorkbookEncryptionHelper.IsWorkbookEncrypted(workbookPath);
                Console.WriteLine($"Workbook \"{workbookPath}\" encrypted: {encrypted}");
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}