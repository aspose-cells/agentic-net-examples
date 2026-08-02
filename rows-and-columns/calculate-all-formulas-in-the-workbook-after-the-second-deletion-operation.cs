// Title: C# – Recalculate all formulas after deleting rows and columns with UpdateReference using Aspose.Cells
// Description: Demonstrates how to delete a specific row and a column while preserving formula references (DeleteOptions.UpdateReference = true), then recalculate every formula in the workbook with Workbook.CalculateFormula() before saving the file.
// Keywords: Aspose.Cells DeleteRows UpdateReference | Aspose.Cells DeleteColumns UpdateReference | Workbook.CalculateFormula C# | recalculate formulas after row deletion | recalculate formulas after column removal | Aspose.Cells formula refresh | C# spreadsheet structural changes | Aspose.Cells .NET example
// Common Searches: Aspose.Cells recalculate formulas after DeleteRows | How to update formula references when deleting rows in Aspose.Cells | C# calculate all formulas after deleting a column with Aspose.Cells | DeleteOptions.UpdateReference example Aspose.Cells | Workbook.CalculateFormula after structural changes
// Developer Intent: Refresh every formula in a workbook after removing a row and a column with reference updates enabled.
// Use Cases: Maintain accurate totals when a data row is removed from a financial report. | Automatically adjust range‑based formulas after deleting a source column in a data‑analysis sheet. | Ensure exported Excel files contain correct calculated values after structural modifications.
// AI Prompts: Generate C# code that deletes row 2 and column A using Aspose.Cells with UpdateReference enabled, then calls Workbook.CalculateFormula() to update all formulas. | Show how to read the updated values of cells that moved after deletions in Aspose.Cells. | Explain why DeleteOptions.UpdateReference is required and why a subsequent Workbook.CalculateFormula() call is necessary.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaCalculationAfterDeletion
{
    // Demonstrates how to delete a specific row and a column while preserving formula references (DeleteOptions.UpdateReference = true), then recalculate every formula in the workbook with Workbook.CalculateFormula() before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet and its cells
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some data
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].PutValue(30);

            // Add formulas that reference the data
            cells["B1"].Formula = "=A1*2";      // 20
            cells["B2"].Formula = "=A2*2";      // 40
            cells["B3"].Formula = "=A3*2";      // 60
            cells["C1"].Formula = "=SUM(B1:B3)"; // 120

            // First deletion: delete the second row (index 1)
            // Use DeleteOptions to update references in formulas
            DeleteOptions deleteOptions = new DeleteOptions { UpdateReference = true };
            cells.DeleteRows(1, 1, deleteOptions);

            // Second deletion: delete column A (index 0)
            // Again update references so formulas adjust to the new column positions
            cells.DeleteColumns(0, 1, deleteOptions);

            // After the second deletion, calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Output the results to verify calculation
            Console.WriteLine("After deletions and calculation:");
            Console.WriteLine($"B1 (now C1) value: {cells["C1"].Value}");
            Console.WriteLine($"B2 (now C2) value: {cells["C2"].Value}");
            Console.WriteLine($"B3 (now C3) value: {cells["C3"].Value}");
            Console.WriteLine($"C1 (now D1) sum result: {cells["D1"].Value}");

            // Save the workbook (lifecycle rule: save)
            workbook.Save("ResultAfterDeletion.xlsx");
        }
    }
}
