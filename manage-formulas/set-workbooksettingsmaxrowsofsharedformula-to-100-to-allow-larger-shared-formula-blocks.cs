// Title: Aspose.Cells for .NET – Set MaxRowsOfSharedFormula to Control Shared Formula Range
// Description: Demonstrates how to use Workbook.Settings.MaxRowsOfSharedFormula in Aspose.Cells to limit a shared formula to 100 rows, detect truncation, and then raise the limit (e.g., to 1024) so the formula spans the full range. Includes verification of the last cell's formula and saving the workbook.
// Keywords: Aspose.Cells | C# | .NET | Workbook.Settings.MaxRowsOfSharedFormula | shared formula limit | increase shared formula rows | Excel shared formula | Aspose.Cells example | GitHub code snippet | formula truncation
// Common Searches: Aspose.Cells MaxRowsOfSharedFormula example | how to change shared formula row limit in .NET | shared formula stops at 100 rows Aspose.Cells | increase shared formula range Aspose.Cells C# | Workbook.Settings.MaxRowsOfSharedFormula usage
// Developer Intent: Adjust the MaxRowsOfSharedFormula property to define how many rows a shared formula may cover, detect when the default limit truncates the formula, and increase the limit for larger datasets.
// Use Cases: Apply a shared formula to a known range without exceeding the default 100‑row limit. | Dynamically raise MaxRowsOfSharedFormula when processing worksheets with more rows than the default. | Validate that a shared formula has been applied to the final row after changing the limit.
// AI Prompts: Provide C# code that sets Workbook.Settings.MaxRowsOfSharedFormula, adds a shared formula across 101 rows, and prints the formula in the last cell. | Generate a function that calculates the required MaxRowsOfSharedFormula value based on a worksheet's row count before applying a shared formula. | Explain performance considerations when increasing MaxRowsOfSharedFormula and how to handle potential memory impact.

using System;
using Aspose.Cells;

// Demonstrates how to use Workbook.Settings.MaxRowsOfSharedFormula in Aspose.Cells to limit a shared formula to 100 rows, detect truncation, and then raise the limit (e.g., to 1024) so the formula spans the full range. Includes verification of the last cell's formula and saving the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the maximum number of rows that a shared formula can cover
        workbook.Settings.MaxRowsOfSharedFormula = 100;

        // Attempt to set a shared formula that spans 101 rows (exceeds the limit)
        // Only the first 100 rows will receive the formula due to the setting above
        workbook.Worksheets[0].Cells["B1"].SetSharedFormula("=A1", 101, 1);

        // Verify the formula in the last cell of the range (B101)
        // It will be empty because the shared formula was truncated at row 100
        Console.WriteLine("Formula in B101 (first sheet): " + workbook.Worksheets[0].Cells["B101"].Formula);

        // Increase the limit to allow the full range of 101 rows
        workbook.Settings.MaxRowsOfSharedFormula = 1024;

        // Add a new worksheet and set the same shared formula
        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
        sheet2.Cells["B1"].SetSharedFormula("=A1", 101, 1);

        // Verify that the formula now exists in B101 of the second sheet
        Console.WriteLine("Formula in B101 (Sheet2): " + sheet2.Cells["B101"].Formula);

        // Save the workbook to a file
        workbook.Save("output.xlsx");
    }
}
