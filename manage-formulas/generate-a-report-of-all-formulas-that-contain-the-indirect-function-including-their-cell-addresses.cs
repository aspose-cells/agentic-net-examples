// Title: Create a .NET C# worksheet that lists all cells containing the INDIRECT function with their addresses using Aspose.Cells
// Common Searches: C# Aspose.Cells find all cells with INDIRECT formula and export addresses | how to generate a report of INDIRECT formulas in an Excel workbook using Aspose.Cells | list cell addresses for formulas that use INDIRECT with Aspose.Cells for .NET | extract and write INDIRECT function formulas to a new sheet in C# | Aspose.Cells enumerate formulas containing INDIRECT and save report workbook
// Tags: list INDIRECT formulas Aspose.Cells | extract cell formulas containing INDIRECT C# | create report worksheet Aspose.Cells .NET | case-insensitive formula search Aspose.Cells | add audit sheet to workbook using Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The example loads an existing Excel file, iterates through every worksheet and cell, detects formulas that include the INDIRECT function (case‑insensitive), records each cell's address and formula, adds a new worksheet named "IndirectFormulaReport", writes the collected data as CSV‑style rows, and saves the workbook with the added report sheet.
class Program
{
    static void Main()
    {
        // Load the existing workbook (load rule)
        var workbook = new Workbook("input.xlsx");

        // Prepare a list to store report lines (header + data)
        var reportLines = new List<string>();
        reportLines.Add("CellAddress,Formula");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Enumerate all cells that contain data/formulas
            foreach (Cell cell in cells)
            {
                // Check if the cell contains a formula
                if (cell.IsFormula)
                {
                    string formula = cell.Formula;

                    // Look for the INDIRECT function (case‑insensitive)
                    if (formula.IndexOf("INDIRECT", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // Add the cell address and its formula to the report
                        reportLines.Add($"{cell.Name},{formula}");
                    }
                }
            }
        }

        // Add a new worksheet to hold the report (create rule)
        int reportSheetIndex = workbook.Worksheets.Add();
        Worksheet reportSheet = workbook.Worksheets[reportSheetIndex];
        reportSheet.Name = "IndirectFormulaReport";

        // Write the report lines into the new worksheet
        for (int i = 0; i < reportLines.Count; i++)
        {
            reportSheet.Cells[i, 0].PutValue(reportLines[i]);
        }

        // Save the workbook with the added report sheet (save rule)
        workbook.Save("output.xlsx");
    }
}
