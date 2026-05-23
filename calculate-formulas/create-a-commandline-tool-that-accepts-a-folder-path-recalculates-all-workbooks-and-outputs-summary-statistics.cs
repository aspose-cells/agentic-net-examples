using System;
using System.IO;
using System.Linq;
using Aspose.Cells;

class Program
{
    static void Main(string[] args)
    {
        try
        {
            // Verify that a folder path was provided
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: RecalcTool <folderPath>");
                return;
            }

            string folderPath = args[0];

            // Ensure the folder exists
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder not found: {folderPath}");
                return;
            }

            // Find Excel files in the folder (common extensions)
            var excelFiles = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly)
                .Where(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                            f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) ||
                            f.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase) ||
                            f.EndsWith(".xlsb", StringComparison.OrdinalIgnoreCase))
                .ToList();

            int totalWorkbooks = 0;
            int totalWorksheets = 0;

            foreach (string filePath in excelFiles)
            {
                // Verify the file still exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found (skipped): {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Recalculate all formulas in the workbook
                    workbook.CalculateFormula();

                    // Save the workbook back to the same file
                    workbook.Save(filePath);

                    totalWorkbooks++;
                    totalWorksheets += workbook.Worksheets.Count;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{Path.GetFileName(filePath)}': {ex.Message}");
                }
            }

            // Output overall statistics
            Console.WriteLine($"Processed {totalWorkbooks} workbook(s) containing a total of {totalWorksheets} worksheet(s).");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}