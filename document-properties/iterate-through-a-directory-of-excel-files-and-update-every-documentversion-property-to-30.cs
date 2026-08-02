using System;
using System.IO;
using System.Linq;
using Aspose.Cells;

namespace UpdateDocumentVersion
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the directory containing Excel files
            string directoryPath = @"C:\Path\To\ExcelFiles";

            // Verify that the directory exists
            if (!Directory.Exists(directoryPath))
            {
                Console.WriteLine($"Directory not found: {directoryPath}");
                return;
            }

            // Define the file extensions to process
            string[] extensions = new[] { ".xls", ".xlsx", ".xlsm", ".ods", ".csv" };

            // Get all files with the specified extensions
            var excelFiles = Directory.GetFiles(directoryPath, "*.*", SearchOption.TopDirectoryOnly)
                                      .Where(f => extensions.Any(ext => f.EndsWith(ext, StringComparison.OrdinalIgnoreCase)))
                                      .ToList();

            foreach (var filePath in excelFiles)
            {
                // Ensure the file still exists before processing
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found (skipped): {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook from the file
                    Workbook workbook = new Workbook(filePath);

                    // Update the built‑in DocumentVersion property
                    workbook.BuiltInDocumentProperties.DocumentVersion = "3.0";

                    // Save the workbook back to the same file
                    workbook.Save(filePath);

                    Console.WriteLine($"Updated DocumentVersion for: {Path.GetFileName(filePath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Processing completed.");
        }
    }
}