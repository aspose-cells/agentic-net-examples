// Title: Export IFERROR formulas to a separate review worksheet with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells extract cells that contain IFERROR and write them to a new worksheet | How to list all IFERROR formulas in an Excel file using Aspose.Cells .NET | Generate a review sheet of IFERROR formulas with Aspose.Cells in C# | Save formulas with IFERROR to a separate sheet programmatically using Aspose.Cells
// Tags: extract IFERROR formulas Aspose.Cells | write formula details to new worksheet C# | scan used range for specific functions Aspose.Cells | create review sheet for error handling formulas .NET | collect cell addresses with IFERROR Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;

// The example loads an existing workbook, scans each worksheet's used range for cells whose formulas contain the IFERROR function, records the sheet name, cell address, and formula, creates a new worksheet named "IFERROR Review", writes the collected information as a table, and saves the updated workbook as a new file.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Collect information about cells that contain IFERROR in their formulas
        var formulaInfos = new List<(string SheetName, string CellAddress, string Formula)>();

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Determine the used range of the worksheet
            int maxRow = sheet.Cells.MaxDataRow;
            int maxColumn = sheet.Cells.MaxDataColumn;

            // Scan each cell within the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell cell = sheet.Cells[row, col];

                    // Check if the cell contains a formula with IFERROR (case‑insensitive)
                    if (cell.IsFormula && cell.Formula.IndexOf("IFERROR", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        formulaInfos.Add((sheet.Name, cell.Name, cell.Formula));
                    }
                }
            }
        }

        // Add a new worksheet to hold the review list
        int reviewSheetIndex = workbook.Worksheets.Add();
        Worksheet reviewSheet = workbook.Worksheets[reviewSheetIndex];
        reviewSheet.Name = "IFERROR Review";

        // Write header row
        reviewSheet.Cells[0, 0].PutValue("Sheet");
        reviewSheet.Cells[0, 1].PutValue("Cell");
        reviewSheet.Cells[0, 2].PutValue("Formula");

        // Populate the review sheet with the collected formula data
        for (int i = 0; i < formulaInfos.Count; i++)
        {
            var info = formulaInfos[i];
            reviewSheet.Cells[i + 1, 0].PutValue(info.SheetName);
            reviewSheet.Cells[i + 1, 1].PutValue(info.CellAddress);
            reviewSheet.Cells[i + 1, 2].PutValue(info.Formula);
        }

        // Save the modified workbook to a new file
        workbook.Save("output.xlsx");
    }
}
