// Title: Export IFERROR formulas from an Excel workbook to a review worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that scans every worksheet, finds cells whose formula contains IFERROR, and writes the sheet name, cell address, and formula to a newly added worksheet called "IFERROR Review". | Update the script to also capture the evaluated result of each IFERROR formula and include it as an additional column in the review sheet.
// Common Searches: aspnet aspocells find IFERROR formulas in workbook | c# extract cells containing IFERROR using Aspose.Cells | how to list error‑handling formulas from an Excel file with Aspose.Cells .NET | save locations of IFERROR formulas to a new sheet in Aspose.Cells | filter workbook formulas by function name Aspose.Cells C#
// Tags: extract IFERROR formulas Aspose.Cells | list formula cells by function .NET | create review worksheet Aspose.Cells | export formula metadata to new sheet C# | filter cells with error handling function Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The program loads an Excel workbook, iterates through all worksheets and cells to locate formulas containing the IFERROR function, records each occurrence's sheet name, cell address, and formula, writes this information to a newly added worksheet named "IFERROR Review", and saves the updated workbook.
class ExportIfErrorFormulas
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // List to hold information about cells containing IFERROR
        List<(string SheetName, string CellName, string Formula)> ifErrorCells = new List<(string, string, string)>();

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all cells that have formulas
            foreach (Cell cell in sheet.Cells)
            {
                if (cell.IsFormula && cell.Formula != null &&
                    cell.Formula.IndexOf("IFERROR", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // Store sheet name, cell name (e.g., A1), and the formula text
                    ifErrorCells.Add((sheet.Name, cell.Name, cell.Formula));
                }
            }
        }

        // Add a new worksheet to hold the review list
        int reviewSheetIndex = workbook.Worksheets.Add();
        Worksheet reviewSheet = workbook.Worksheets[reviewSheetIndex];
        reviewSheet.Name = "IFERROR Review";

        // Write header row
        reviewSheet.Cells["A1"].PutValue("Worksheet");
        reviewSheet.Cells["B1"].PutValue("Cell");
        reviewSheet.Cells["C1"].PutValue("Formula");

        // Populate the review sheet with collected data
        int currentRow = 2; // Start after header
        foreach (var entry in ifErrorCells)
        {
            reviewSheet.Cells[currentRow, 0].PutValue(entry.SheetName);   // Column A
            reviewSheet.Cells[currentRow, 1].PutValue(entry.CellName);   // Column B
            reviewSheet.Cells[currentRow, 2].PutValue(entry.Formula);    // Column C
            currentRow++;
        }

        // Save the workbook with the new worksheet (replace with desired output path)
        workbook.Save("output.xlsx");
    }
}
