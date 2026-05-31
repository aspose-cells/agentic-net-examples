using System;
using System.IO;
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

                // Verify folder exists
                if (!Directory.Exists(folderPath))
                {
                    Console.WriteLine($"Folder not found: {folderPath}");
                    return;
                }

                // Retrieve all files in the folder
                string[] allFiles = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly);

                foreach (string filePath in allFiles)
                {
                    try
                    {
                        // Process only supported Excel formats
                        string ext = Path.GetExtension(filePath).ToLowerInvariant();
                        if (ext != ".xls" && ext != ".xlsx" && ext != ".xlsm" && ext != ".xlsb" && ext != ".csv")
                            continue;

                        // Ensure the file exists before loading
                        if (!File.Exists(filePath))
                        {
                            Console.WriteLine($"File not found: {filePath}");
                            continue;
                        }

                        // Load the workbook (lifecycle rule: load)
                        Workbook workbook = new Workbook(filePath);

                        // Set the built‑in document property "Language" to "en-GB"
                        workbook.BuiltInDocumentProperties.Language = "en-GB";

                        // Save the workbook back to the same file (lifecycle rule: save)
                        workbook.Save(filePath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}