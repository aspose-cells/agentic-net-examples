// Title: C# command‑line tool to batch‑recalculate formulas in all Excel workbooks within a folder and display processing summary using Aspose.Cells
// Common Searches: C# program to recalculate formulas in every Excel file in a directory with Aspose.Cells | How to batch process Excel workbooks and count worksheets using Aspose.Cells in a console app | Command line utility for recalculating and saving multiple .xlsx files with Aspose.Cells | Get total number of formula cells across all workbooks in a folder using Aspose.Cells C# | Aspose.Cells calculate formulas for all files in a folder and output processing summary
// Tags: batch recalculate formulas Aspose.Cells C# | enumerate Excel workbooks in folder Aspose.Cells | count worksheets and formula cells Aspose.Cells | overwrite original Excel files after calculation Aspose.Cells | console application summary statistics Excel processing | process .xls .xlsx .xlsm .xlsb files Aspose.Cells

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;

namespace RecalculateWorkbooks
{
    // The console application accepts a folder path, scans for .xls, .xlsx, .xlsm, and .xlsb files, loads each workbook with Aspose.Cells, counts its worksheets and formula cells, recalculates all formulas, overwrites the original file, and logs per‑file details. After processing, it prints total workbooks, total worksheets, and total formula cells processed.
    class Program
    {
        static void Main(string[] args)
        {
            // Validate arguments
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: RecalculateWorkbooks <folderPath>");
                return;
            }

            string folderPath = args[0];

            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Error: Folder \"{folderPath}\" does not exist.");
                return;
            }

            // Supported Excel extensions
            string[] extensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb" };

            // Gather all workbook files in the folder (non‑recursive)
            var workbookFiles = Directory.GetFiles(folderPath)
                                         .Where(f => extensions.Contains(Path.GetExtension(f), StringComparer.OrdinalIgnoreCase))
                                         .ToList();

            if (!workbookFiles.Any())
            {
                Console.WriteLine("No Excel workbooks found in the specified folder.");
                return;
            }

            // Summary statistics
            int totalWorkbooks = 0;
            int totalWorksheets = 0;
            int totalFormulaCells = 0;

            foreach (var filePath in workbookFiles)
            {
                try
                {
                    // Load workbook
                    Workbook workbook = new Workbook(filePath);

                    // Count worksheets
                    int worksheetCount = workbook.Worksheets.Count;
                    totalWorksheets += worksheetCount;

                    // Count formula cells before recalculation
                    int formulaCount = 0;
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Get all cells with formulas
                        var cells = sheet.Cells;
                        foreach (Cell cell in cells)
                        {
                            if (cell.IsFormula)
                                formulaCount++;
                        }
                    }
                    totalFormulaCells += formulaCount;

                    // Recalculate all formulas
                    workbook.CalculateFormula();

                    // Save the workbook (overwrite original)
                    workbook.Save(filePath);

                    totalWorkbooks++;
                    Console.WriteLine($"Processed: {Path.GetFileName(filePath)} (Worksheets: {worksheetCount}, Formulas: {formulaCount})");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to process \"{Path.GetFileName(filePath)}\": {ex.Message}");
                }
            }

            // Output summary
            Console.WriteLine();
            Console.WriteLine("=== Summary ===");
            Console.WriteLine($"Total workbooks processed : {totalWorkbooks}");
            Console.WriteLine($"Total worksheets          : {totalWorksheets}");
            Console.WriteLine($"Total formula cells found : {totalFormulaCells}");
        }
    }
}
