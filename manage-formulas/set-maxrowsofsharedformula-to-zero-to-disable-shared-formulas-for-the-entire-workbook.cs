// Title: Disable shared formulas for an entire workbook using MaxRowsOfSharedFormula = 0 in Aspose.Cells for .NET
// Description: Shows how to turn off shared formulas across a workbook by setting Workbook.Settings.MaxRowsOfSharedFormula to 0. The sample creates a workbook, adds data, applies a formula with SetSharedFormula (which behaves as a normal formula when sharing is disabled), recalculates, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | disable shared formulas | MaxRowsOfSharedFormula | formula settings | Workbook.Settings | SetSharedFormula | Excel compatibility | performance optimization
// Common Searches: Aspose.Cells disable shared formulas | MaxRowsOfSharedFormula 0 effect | turn off shared formulas .NET | prevent shared formula generation Aspose | how to set MaxRowsOfSharedFormula in C#
// Developer Intent: Turn off shared formulas for the whole workbook.
// Use Cases: Maintain compatibility with legacy Excel versions that do not support shared formulas | Facilitate per‑cell formula auditing and debugging | Avoid shared‑formula limits when using custom or volatile functions | Improve stability when exporting large datasets to Excel
// AI Prompts: Generate C# code to re‑enable shared formulas after they have been disabled with MaxRowsOfSharedFormula. | Show how to detect if a workbook has shared formulas disabled using Aspose.Cells. | Compare performance of a workbook with MaxRowsOfSharedFormula set to 0 versus a positive value. | Explain how MaxRowsOfSharedFormula interacts with SetSharedFormula and CalculateFormula methods.

using System;
using Aspose.Cells;

// Shows how to turn off shared formulas across a workbook by setting Workbook.Settings.MaxRowsOfSharedFormula to 0. The sample creates a workbook, adds data, applies a formula with SetSharedFormula (which behaves as a normal formula when sharing is disabled), recalculates, and saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Disable shared formulas for the entire workbook
        workbook.Settings.MaxRowsOfSharedFormula = 0;

        // Example data to show that formulas still work (they just won't be shared)
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate column A with sample values
        for (int i = 0; i < 5; i++)
        {
            cells[i, 0].PutValue(i + 1); // A1..A5 = 1..5
        }

        // Attempt to set a shared formula; with MaxRowsOfSharedFormula = 0 it behaves as a normal formula
        cells["B1"].SetSharedFormula("=A1*2", 5, 1);

        // Calculate formulas to obtain results
        workbook.CalculateFormula();

        // Save the workbook
        workbook.Save("DisabledSharedFormulas.xlsx");
    }
}
