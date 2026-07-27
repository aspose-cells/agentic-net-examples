using System;
using System.IO;
using Aspose.Cells;

namespace BatchWorkbookLanguageUpdater
{
    class Program
    {
        static void Main()
        {
            // Folder containing the workbooks to process
            string folderPath = @"C:\Workbooks";

            // Define the file extensions to consider as Excel workbooks
            string[] extensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb", ".csv" };

            // Iterate through each file in the folder that matches the extensions
            foreach (string filePath in Directory.GetFiles(folderPath))
            {
                // Skip files that are not Excel workbooks
                if (Array.IndexOf(extensions, Path.GetExtension(filePath).ToLower()) < 0)
                    continue;

                // Ensure the file still exists before attempting to load
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook (will throw if the file is password‑protected)
                    using (Workbook workbook = new Workbook(filePath))
                    {
                        // Set the built‑in document property "Language" to "en-GB"
                        workbook.BuiltInDocumentProperties.Language = "en-GB";

                        // Save the workbook back to the same file (overwrites the original)
                        workbook.Save(filePath);
                    }
                }
                catch (CellsException ex) when (ex.Message.Contains("password", StringComparison.OrdinalIgnoreCase))
                {
                    // Skip password‑protected files
                    Console.WriteLine($"Skipping password‑protected file: {filePath}");
                }
                catch (Exception ex)
                {
                    // Log any other errors and continue processing remaining files
                    Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("All applicable workbooks have been processed.");
        }
    }
}