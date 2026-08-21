// Title: Aspose.Cells C# – Expand MaxRowsOfSharedFormula for Larger Shared Formula Ranges
// Description: Demonstrates how to change Workbook.Settings.MaxRowsOfSharedFormula, test the default 100‑row limit, raise it to 1024, apply a shared formula across 101 rows on a new sheet, verify the last cell's formula, and save the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells MaxRowsOfSharedFormula | shared formula row limit .NET | increase shared formula block size | Workbook.Settings MaxRowsOfSharedFormula C# | Aspose.Cells shared formula example
// Common Searches: set MaxRowsOfSharedFormula Aspose.Cells | increase shared formula rows Aspose.Cells .NET | shared formula limit 100 rows Aspose | how to expand shared formula range Aspose.Cells
// Developer Intent: Adjust the maximum row count for shared formula blocks in an Aspose.Cells workbook.
// Use Cases: Configure MaxRowsOfSharedFormula before creating a shared formula to avoid truncated ranges. | Detect when a shared formula exceeds the current limit, raise the limit, and reapply the formula. | Confirm that the formula appears in the final cell after increasing the limit and then persist the workbook.
// AI Prompts: Write C# code that sets MaxRowsOfSharedFormula to 500 and creates a shared formula spanning 300 rows with Aspose.Cells. | Explain how to read the current MaxRowsOfSharedFormula value and increase it only when a shared formula would exceed that value. | Provide a step‑by‑step guide to verify a shared formula after raising MaxRowsOfSharedFormula and then saving the workbook.

using System;
using Aspose.Cells;

// Demonstrates how to change Workbook.Settings.MaxRowsOfSharedFormula, test the default 100‑row limit, raise it to 1024, apply a shared formula across 101 rows on a new sheet, verify the last cell's formula, and save the workbook using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the maximum number of rows allowed for a shared formula block
        workbook.Settings.MaxRowsOfSharedFormula = 100;

        // Attempt to set a shared formula that exceeds the current limit (101 rows)
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;
        cells["B1"].SetSharedFormula("=A1", 101, 1);

        // Verify the formula in the last cell of the range (will be empty because of the limit)
        Console.WriteLine("Formula in B101 (original limit): " + cells["B101"].Formula);

        // Increase the limit to allow larger shared formula blocks
        workbook.Settings.MaxRowsOfSharedFormula = 1024;

        // Set the same shared formula on a new worksheet now that the limit is higher
        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
        Cells cells2 = sheet2.Cells;
        cells2["B1"].SetSharedFormula("=A1", 101, 1);
        Console.WriteLine("Formula in B101 (increased limit): " + cells2["B101"].Formula);

        // Save the workbook to a file
        workbook.Save("output.xlsx");
    }
}
