// Title: Restrict Shared Formula to 50 Rows with Workbook.Settings.MaxRowsOfSharedFormula (Aspose.Cells .NET)
// Description: Shows how to set Workbook.Settings.MaxRowsOfSharedFormula to 50, apply a shared formula across 60 rows, and confirm that only the first 50 rows keep the formula before saving the file.
// Keywords: Aspose.Cells | .NET | C# | MaxRowsOfSharedFormula | shared formula limit | Workbook.Settings | SetSharedFormula | performance optimization | formula scope control
// Common Searches: Aspose.Cells limit shared formula rows | C# MaxRowsOfSharedFormula property example | How to restrict shared formula range in Aspose.Cells | Set maximum rows for shared formulas .NET | Workbook.Settings.MaxRowsOfSharedFormula usage
// Developer Intent: Configure the workbook so that shared formulas are applied to no more than 50 rows.
// Use Cases: Avoid performance hits in large sheets by capping shared‑formula propagation. | Define a fixed calculation range for financial or statistical reports. | Prevent accidental formula evaluation in rows beyond a required boundary.
// AI Prompts: Generate code that checks whether rows after the MaxRowsOfSharedFormula limit contain any formula. | Create a routine that adjusts MaxRowsOfSharedFormula based on the worksheet's data size. | Explain what happens when SetSharedFormula requests more rows than the MaxRowsOfSharedFormula setting allows.

using System;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // Shows how to set Workbook.Settings.MaxRowsOfSharedFormula to 50, apply a shared formula across 60 rows, and confirm that only the first 50 rows keep the formula before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Limit the maximum number of rows that a shared formula can span to 50
            workbook.Settings.MaxRowsOfSharedFormula = 50;

            // Example: set a shared formula that tries to use 60 rows (will be limited by the setting)
            // This will demonstrate that only the first 50 rows are treated as shared formula.
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;
            cells["B1"].SetSharedFormula("=A1", 60, 1); // 60 rows requested

            // Verify the setting value
            Console.WriteLine("MaxRowsOfSharedFormula is set to: " + workbook.Settings.MaxRowsOfSharedFormula);
            Console.WriteLine("Formula in B50: " + cells["B50"].Formula);
            Console.WriteLine("Formula in B51 (should be empty or default): " + cells["B51"].Formula);

            // Save the workbook (lifecycle: save)
            workbook.Save("SharedFormulaLimited.xlsx");
        }
    }
}
