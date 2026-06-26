using System;
using System.IO;
using System.Linq;
using Aspose.Cells;

namespace BatchSetLanguageApp
{
    class BatchSetLanguage
    {
        static void Main()
        {
            try
            {
                // Folder containing the workbooks to process
                string folderPath = @"C:\Path\To\Folder";

                if (!Directory.Exists(folderPath))
                {
                    Console.WriteLine($"Folder not found: {folderPath}");
                    return;
                }

                // Define the file extensions to consider as Excel workbooks
                string[] extensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb", ".csv" };

                // Get all matching files in the folder (non‑recursive)
                var workbookFiles = Directory.GetFiles(folderPath)
                    .Where(f => extensions.Any(ext => f.EndsWith(ext, StringComparison.OrdinalIgnoreCase)))
                    .ToArray();

                foreach (var filePath in workbookFiles)
                {
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found, skipping: {filePath}");
                        continue;
                    }

                    try
                    {
                        // Load the workbook from the file
                        Workbook workbook = new Workbook(filePath);

                        // Set the built‑in document property "Language" to en‑GB
                        workbook.BuiltInDocumentProperties.Language = "en-GB";

                        // Save the workbook, overwriting the original file
                        workbook.Save(filePath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                    }
                }

                Console.WriteLine("Language property set to en-GB for all workbooks in the folder.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}