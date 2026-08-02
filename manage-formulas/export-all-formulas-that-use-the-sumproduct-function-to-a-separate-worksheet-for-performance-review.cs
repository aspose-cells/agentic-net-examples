// Title: Export SUMPRODUCT formulas to a review worksheet with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, adds a sheet called “SUMPRODUCT Review”, scans every worksheet for formulas containing the SUMPRODUCT function (case‑insensitive), records the source sheet name, cell address and full formula, then saves the file for performance analysis.
// Keywords: Aspose.Cells | C# | .NET | SUMPRODUCT | extract formulas | export formulas to new sheet | Excel performance review | formula audit | iterate workbook cells | diagnostic worksheet
// Common Searches: Aspose.Cells export SUMPRODUCT formulas C# | How to list all SUMPRODUCT functions in an Excel file using .NET | Create a review sheet for specific Excel formulas with Aspose.Cells | Iterate through workbook cells and collect formulas containing a keyword | Performance analysis of SUMPRODUCT formulas in C#
// Developer Intent: Extract every SUMPRODUCT formula and write its details to a separate worksheet for analysis.
// Use Cases: Generate a report of all SUMPRODUCT formulas to identify calculation bottlenecks. | Provide a diagnostic sheet for auditors to review complex array formulas. | Create documentation of SUMPRODUCT usage across multiple worksheets for knowledge transfer.
// AI Prompts: Write C# code with Aspose.Cells that extracts formulas containing a specified function and logs them to a new worksheet. | Modify the example to also capture each SUMPRODUCT formula's evaluated value in the review sheet. | Add logic to ignore hidden worksheets and include a summary count of SUMPRODUCT formulas per sheet.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an Excel workbook, adds a sheet called “SUMPRODUCT Review”, scans every worksheet for formulas containing the SUMPRODUCT function (case‑insensitive), records the source sheet name, cell address and full formula, then saves the file for performance analysis.
    public class ExportSumProductFormulas
    {
        public static void Run()
        {
            try
            {
                // Path to the source workbook
                string sourcePath = "input.xlsx";

                // Verify source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file '{sourcePath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(sourcePath);

                // Add a new worksheet to hold the extracted SUMPRODUCT formulas
                int newSheetIndex = workbook.Worksheets.Add();
                Worksheet reviewSheet = workbook.Worksheets[newSheetIndex];
                reviewSheet.Name = "SUMPRODUCT Review";

                // Write header row in the review sheet
                Cells reviewCells = reviewSheet.Cells;
                reviewCells["A1"].PutValue("Source Sheet");
                reviewCells["B1"].PutValue("Cell Address");
                reviewCells["C1"].PutValue("Formula");

                int reviewRowIndex = 1; // start after header (0‑based index)

                // Iterate through all worksheets and cells
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    // Skip the review sheet itself
                    if (ws.Name == reviewSheet.Name) continue;

                    Cells cells = ws.Cells;
                    int maxRow = cells.MaxDataRow;
                    int maxCol = cells.MaxDataColumn;

                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = cells[row, col];
                            if (!string.IsNullOrEmpty(cell.Formula))
                            {
                                // Look for the SUMPRODUCT function (case‑insensitive)
                                if (cell.Formula.IndexOf("SUMPRODUCT", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    reviewCells[reviewRowIndex, 0].PutValue(ws.Name);      // Source sheet name
                                    reviewCells[reviewRowIndex, 1].PutValue(cell.Name);   // Cell address (e.g., B5)
                                    reviewCells[reviewRowIndex, 2].PutValue(cell.Formula); // Full formula
                                    reviewRowIndex++;
                                }
                            }
                        }
                    }
                }

                // Save the modified workbook
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Export completed. Review sheet saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportSumProductFormulas.Run();
        }
    }
}
