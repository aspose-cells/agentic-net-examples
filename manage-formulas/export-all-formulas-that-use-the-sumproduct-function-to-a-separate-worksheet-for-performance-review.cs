// Title: Export SUMPRODUCT formulas from an Excel workbook to a dedicated review worksheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that iterates all worksheets, detects cells whose formula contains the SUMPRODUCT function, and writes the source sheet name, cell address, and formula to a newly added 'SUMPRODUCT Review' sheet. | Create a method that adds a review worksheet, records each SUMPRODUCT formula found, auto‑fits the columns for readability, and saves the workbook as a new file.
// Common Searches: how to list cells with SUMPRODUCT formula using Aspose.Cells C# | Aspose.Cells extract specific formulas to another worksheet .NET | C# code to export SUMPRODUCT calculations from Excel for performance analysis | scan Excel workbook for SUMPRODUCT functions and generate report with Aspose.Cells | auto‑fit columns after writing data to a new sheet in Aspose.Cells
// Tags: export SUMPRODUCT formulas Aspose.Cells | create formula review sheet C# Aspose.Cells | scan workbook for specific formula Aspose.Cells | auto‑fit columns after data export Aspose.Cells | save modified workbook C# Aspose.Cells

using System;
using Aspose.Cells;

// The program loads an existing workbook, adds a worksheet named "SUMPRODUCT Review", iterates through all other worksheets to locate cells whose formulas contain the SUMPRODUCT function, records the source sheet name, cell address, and formula in the review sheet, auto‑fits the columns for readability, and saves the updated workbook as a new file.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Add a new worksheet to hold SUMPRODUCT formulas
        int newSheetIndex = workbook.Worksheets.Add();
        Worksheet reviewSheet = workbook.Worksheets[newSheetIndex];
        reviewSheet.Name = "SUMPRODUCT Review";

        // Write header row
        reviewSheet.Cells[0, 0].PutValue("Source Sheet");
        reviewSheet.Cells[0, 1].PutValue("Cell Address");
        reviewSheet.Cells[0, 2].PutValue("Formula");

        int outputRow = 1; // Start after header

        // Iterate through all worksheets except the review sheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            if (sheet.Name == reviewSheet.Name)
                continue;

            // Determine the used range
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    if (cell.IsFormula)
                    {
                        string formula = cell.Formula;
                        if (!string.IsNullOrEmpty(formula) &&
                            formula.IndexOf("SUMPRODUCT", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Record the formula details in the review sheet
                            reviewSheet.Cells[outputRow, 0].PutValue(sheet.Name);
                            reviewSheet.Cells[outputRow, 1].PutValue(cell.Name); // e.g., A1
                            reviewSheet.Cells[outputRow, 2].PutValue(formula);
                            outputRow++;
                        }
                    }
                }
            }
        }

        // Adjust column widths for readability
        reviewSheet.AutoFitColumns();

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
