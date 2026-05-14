using System;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsBatchProcessing
{
    public static class LanguageUpdater
    {
        // Updates the Language built‑in document property of every workbook in the specified folder to "en-GB".
        public static void ProcessFolder(string folderPath)
        {
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder does not exist: {folderPath}");
                return;
            }

            // Define the file extensions that represent Excel workbooks.
            string[] extensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb", ".csv" };

            // Enumerate all matching files in the folder (non‑recursive).
            var workbookFiles = Directory.GetFiles(folderPath)
                                         .Where(f => extensions.Contains(Path.GetExtension(f), StringComparer.OrdinalIgnoreCase));

            foreach (var filePath in workbookFiles)
            {
                try
                {
                    // Load the workbook using the standard constructor (create/load rule).
                    Workbook workbook = new Workbook(filePath);

                    // Set the Language property to "en-GB" (built‑in document property rule).
                    workbook.BuiltInDocumentProperties.Language = "en-GB";

                    // Save the workbook back to the same file (save rule).
                    workbook.Save(filePath);

                    Console.WriteLine($"Updated language for: {Path.GetFileName(filePath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to process {Path.GetFileName(filePath)}: {ex.Message}");
                }
            }
        }

        // Example entry point.
        public static void Main(string[] args)
        {
            // Expect the folder path as the first argument; otherwise use the current directory.
            string targetFolder = args.Length > 0 ? args[0] : Directory.GetCurrentDirectory();

            ProcessFolder(targetFolder);
        }
    }
}