// Title: Export SUMPRODUCT formulas to a review worksheet using Aspose.Cells for .NET (C#)
// Description: Loads an existing workbook, creates a sheet called "SUMPRODUCT Review", writes headers, scans every worksheet’s used range, detects formulas that contain the SUMPRODUCT function (case‑insensitive), records the source sheet name, cell address, and formula, and saves the workbook as a new file for performance analysis.
// Keywords: Aspose.Cells | C# | .NET | SUMPRODUCT | export formulas | review worksheet | Excel performance | extract formulas | list formulas | used range iteration | save workbook
// Common Searches: Aspose.Cells find SUMPRODUCT formulas in Excel | Export SUMPRODUCT formulas to another sheet C# | How to list cells with SUMPRODUCT using Aspose.Cells | Create a review worksheet for specific formulas .NET | Iterate used range of worksheets Aspose.Cells
// Developer Intent: Locate every cell that uses the SUMPRODUCT function and log its sheet, address, and formula on a separate worksheet for analysis or optimization.
// Use Cases: Generate a performance‑review report of all SUMPRODUCT calculations across a workbook. | Provide auditors with a concise list of heavy formulas for optimization assessment. | Automate documentation of formula usage for compliance or knowledge‑base creation.
// AI Prompts: Write C# code with Aspose.Cells that scans all worksheets, finds cells containing SUMPRODUCT, and writes the sheet name, cell address, and formula to a new worksheet. | Modify the example to ignore hidden worksheets while exporting SUMPRODUCT formulas. | Explain how to extend the solution to capture SUMPRODUCT when it appears inside other functions such as IF(SUMPRODUCT(...)).

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an existing workbook, creates a sheet called "SUMPRODUCT Review", writes headers, scans every worksheet’s used range, detects formulas that contain the SUMPRODUCT function (case‑insensitive), records the source sheet name, cell address, and formula, and saves the workbook as a new file for performance analysis.
    public class ExportSumProductFormulas
    {
        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the source workbook
                Workbook workbook = new Workbook(inputPath);

                // Add a new worksheet to hold the extracted SUMPRODUCT formulas
                Worksheet reviewSheet = workbook.Worksheets[workbook.Worksheets.Add()];
                reviewSheet.Name = "SUMPRODUCT Review";

                // Write header row in the review sheet
                reviewSheet.Cells[0, 0].PutValue("Source Sheet");
                reviewSheet.Cells[0, 1].PutValue("Cell Address");
                reviewSheet.Cells[0, 2].PutValue("Formula");

                int reviewRow = 1; // Start writing data from the second row

                // Iterate through all worksheets in the workbook
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    // Skip the review sheet itself (if it already existed)
                    if (ws.Name == reviewSheet.Name) continue;

                    // Determine the used range to limit the iteration
                    int maxRow = ws.Cells.MaxDataRow;
                    int maxCol = ws.Cells.MaxDataColumn;

                    // Scan each cell within the used range
                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = ws.Cells[row, col];

                            // Check if the cell contains a formula that uses SUMPRODUCT
                            if (cell.IsFormula && !string.IsNullOrEmpty(cell.Formula) &&
                                cell.Formula.IndexOf("SUMPRODUCT", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                // Record the information in the review sheet
                                reviewSheet.Cells[reviewRow, 0].PutValue(ws.Name);
                                reviewSheet.Cells[reviewRow, 1].PutValue(cell.Name); // e.g., "B2"
                                reviewSheet.Cells[reviewRow, 2].PutValue(cell.Formula);
                                reviewRow++;
                            }
                        }
                    }
                }

                // Save the modified workbook with the new review worksheet
                workbook.Save(outputPath);
                Console.WriteLine($"Review worksheet saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportSumProductFormulas.Run();
        }
    }
}
