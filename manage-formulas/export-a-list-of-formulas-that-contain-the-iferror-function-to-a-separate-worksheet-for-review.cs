// Title: Export IFERROR formulas from an Excel workbook to a separate review worksheet using Aspose.Cells for .NET (C#)
// Common Searches: how to list cells with IFERROR formula using Aspose.Cells in C# | Aspose.Cells C# extract formulas that contain IFERROR to a separate sheet | create a review worksheet for IFERROR formulas in an Excel file with Aspose.Cells | save IFERROR formula text and cell address programmatically with Aspose.Cells .NET | iterate through all worksheets and collect IFERROR formulas using Aspose.Cells API
// Tags: IFERROR formula extraction Aspose.Cells | programmatic review worksheet creation Aspose.Cells | cell iteration for formula analysis C# | export formula metadata to new Excel sheet .NET | error handling formula collection Excel

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The code loads an Excel workbook, scans every worksheet for cells whose formula contains the IFERROR function (case‑insensitive), records each sheet name, cell address, and formula, creates a new worksheet named "IFERROR Review", writes the collected data with headers, and saves the updated workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        Workbook workbook = null;

        // Load the workbook safely
        try
        {
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                workbook = new Workbook();
            }
        }
        catch (CellsException ex)
        {
            // Handle password‑protected or corrupted files
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            Console.WriteLine("Creating a new empty workbook instead.");
            workbook = new Workbook();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error while loading workbook: {ex.Message}");
            return;
        }

        try
        {
            // Collect information about cells that contain IFERROR in their formulas
            var ifErrorFormulas = new List<(string SheetName, string CellName, string Formula)>();

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Iterate through all cells in the worksheet
                foreach (Cell cell in cells)
                {
                    // Check if the cell has a formula and if it contains IFERROR (case‑insensitive)
                    if (cell.IsFormula && cell.Formula.IndexOf("IFERROR", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        ifErrorFormulas.Add((sheet.Name, cell.Name, cell.Formula));
                    }
                }
            }

            // Add a new worksheet to hold the review list
            int newSheetIndex = workbook.Worksheets.Add();
            Worksheet reviewSheet = workbook.Worksheets[newSheetIndex];
            reviewSheet.Name = "IFERROR Review";

            // Write header row
            reviewSheet.Cells["A1"].PutValue("Sheet");
            reviewSheet.Cells["B1"].PutValue("Cell");
            reviewSheet.Cells["C1"].PutValue("Formula");

            // Write each collected formula to the new worksheet
            int rowIndex = 1; // zero‑based index; row 1 is the second row (after header)
            foreach (var entry in ifErrorFormulas)
            {
                reviewSheet.Cells[rowIndex, 0].PutValue(entry.SheetName);
                reviewSheet.Cells[rowIndex, 1].PutValue(entry.CellName);
                reviewSheet.Cells[rowIndex, 2].PutValue(entry.Formula);
                rowIndex++;
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during processing: {ex.Message}");
        }
    }
}
