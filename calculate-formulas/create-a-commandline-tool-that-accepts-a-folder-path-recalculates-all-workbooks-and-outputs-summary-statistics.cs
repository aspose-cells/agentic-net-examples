// Title: C# command‑line tool to recalculate all formulas in Excel workbooks within a folder and display processing statistics using Aspose.Cells
// Common Searches: c# aspocells recalculate all formulas in multiple Excel files from a folder | how to count formula cells while batch processing workbooks with Aspose.Cells | command line utility to update formulas and get worksheet statistics in .xlsx files using Aspose.Cells | process all Excel files in a directory and recalculate formulas with Aspose.Cells C# | batch recalculate and save Excel workbooks from a folder using Aspose.Cells .NET
// Tags: batch workbook formula recalculation Aspose.Cells | count formula cells Aspose.Cells C# | process Excel files in directory Aspose.Cells | aggregate worksheet statistics Aspose.Cells | save recalculated workbooks Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace RecalculateWorkbooks
{
    // A C# console application that accepts a folder path, scans for supported Excel and CSV files, loads each workbook with Aspose.Cells, counts cells containing formulas, recalculates all formulas, overwrites the original files, and aggregates statistics such as number of workbooks processed, total worksheets, and total formula cells before printing a summary to the console.
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
                Console.WriteLine($"Error: The folder \"{folderPath}\" does not exist.");
                return;
            }

            // Supported workbook extensions
            string[] extensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb", ".csv" };

            // Statistics variables
            int filesProcessed = 0;
            int totalWorksheets = 0;
            long totalFormulaCells = 0;

            // Enumerate all files with supported extensions
            foreach (string filePath in Directory.EnumerateFiles(folderPath, "*.*", SearchOption.AllDirectories))
            {
                if (Array.IndexOf(extensions, Path.GetExtension(filePath).ToLowerInvariant()) < 0)
                    continue; // Skip unsupported files

                // Ensure the file actually exists before attempting to load
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Skipped missing file: {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Count formula cells before recalculation
                    long formulaCountInWorkbook = 0;
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Iterate through all cells in the worksheet and count those containing formulas
                        foreach (Cell cell in sheet.Cells)
                        {
                            if (cell.IsFormula)
                                formulaCountInWorkbook++;
                        }
                    }

                    // Recalculate all formulas in the workbook
                    workbook.CalculateFormula();

                    // Save the workbook back (overwrites original)
                    workbook.Save(filePath);

                    // Update statistics
                    filesProcessed++;
                    totalWorksheets += workbook.Worksheets.Count;
                    totalFormulaCells += formulaCountInWorkbook;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to process \"{filePath}\": {ex.Message}");
                }
            }

            // Output summary statistics
            Console.WriteLine("=== Recalculation Summary ===");
            Console.WriteLine($"Folder processed: {folderPath}");
            Console.WriteLine($"Workbooks processed: {filesProcessed}");
            Console.WriteLine($"Total worksheets: {totalWorksheets}");
            Console.WriteLine($"Total formula cells recalculated: {totalFormulaCells}");
        }
    }
}
