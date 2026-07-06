using System;
using System.IO;
using Aspose.Cells;

namespace RecalculateWorkbooks
{
    class Program
    {
        static void Main(string[] args)
        {
            // Verify that a folder path was provided
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: RecalculateWorkbooks <folderPath>");
                return;
            }

            string folderPath = args[0];

            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder does not exist: {folderPath}");
                return;
            }

            // Supported Excel file extensions
            string[] extensions = new[] { "*.xls", "*.xlsx", "*.xlsm", "*.xlsb" };
            var files = new System.Collections.Generic.List<string>();

            foreach (var ext in extensions)
            {
                files.AddRange(Directory.GetFiles(folderPath, ext, SearchOption.AllDirectories));
            }

            int totalWorkbooks = 0;
            int totalWorksheets = 0;

            foreach (var file in files)
            {
                try
                {
                    // Load the workbook (create rule)
                    Workbook workbook = new Workbook(file);

                    // Recalculate all formulas in the workbook
                    workbook.CalculateFormula();

                    // Save the workbook back to the same file (save rule)
                    workbook.Save(file);

                    // Gather simple statistics
                    int sheetCount = workbook.Worksheets.Count;
                    totalWorkbooks++;
                    totalWorksheets += sheetCount;

                    Console.WriteLine($"Processed: {Path.GetFileName(file)} | Worksheets: {sheetCount}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{file}': {ex.Message}");
                }
            }

            // Output summary statistics
            Console.WriteLine();
            Console.WriteLine("=== Summary ===");
            Console.WriteLine($"Total workbooks processed: {totalWorkbooks}");
            Console.WriteLine($"Total worksheets across all workbooks: {totalWorksheets}");
        }
    }
}