// Title: C# – Update a Cell Formula and Recalculate Workbook with Aspose.Cells
// Description: Loads an Excel file, changes the formula in cell A1 from =B1 to =C1, recalculates all formulas, prints the new value, and saves the workbook as a new file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# formula update | modify Excel formula programmatically | recalculate workbook Aspose | save workbook after formula change | .NET Excel automation | cell reference replacement
// Common Searches: Aspose.Cells change cell formula C# | recalculate formulas after editing Aspose.Cells | how to update Excel formula programmatically .NET | save workbook after formula modification Aspose | verify formula result Aspose.Cells
// Developer Intent: Replace the formula in a specific cell, trigger a full workbook calculation, and persist the updated file.
// Use Cases: Adjust formulas automatically when columns are inserted or removed. | Batch‑replace outdated cell references before generating financial reports. | Validate that a formula edit produces the expected result prior to distribution. | Create dynamic templates where formulas are re‑pointed based on user input.
// AI Prompts: Write C# code using Aspose.Cells to change the formula in cell B2 from "=D5" to "=E5", recalculate only dependent cells, and display the new value. | Explain how to iterate through a worksheet and update multiple formulas to new cell references while ensuring the entire workbook is recalculated with Aspose.Cells for .NET. | Generate a script that loads an Excel workbook, replaces all occurrences of "=SUM(A1:A10)" with "=SUM(B1:B10)", runs CalculateFormula, and saves the result.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaUpdateDemo
{
    // Loads an Excel file, changes the formula in cell A1 from =B1 to =C1, recalculates all formulas, prints the new value, and saves the workbook as a new file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Assume cell A1 contains a formula that originally references B1, e.g., "=B1"
            // Modify the formula to reference C1 instead
            Cell targetCell = cells["A1"];
            // Option 1: directly set the formula string
            targetCell.Formula = "=C1";
            // Option 2: use SetFormula (both achieve the same result)
            // targetCell.SetFormula("=C1", null);

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Verify the updated value of A1 after recalculation
            Console.WriteLine("Updated value in A1: " + targetCell.Value);

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}
