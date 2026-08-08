// Title: Recalculate all workbook formulas with default settings using Workbook.CalculateFormula in Aspose.Cells for .NET
// Description: This C# example creates a workbook, writes a value to A1, adds dependent formulas to B1 and C1, then calls Workbook.CalculateFormula() with no parameters to evaluate every formula using Aspose.Cells' built‑in default calculation engine. The results are printed and the workbook is saved.
// Keywords: Aspose.Cells | .NET | Workbook.CalculateFormula | default calculation engine | evaluate Excel formulas | recalculate workbook | C# formula evaluation | dependent formulas Aspose | Excel calculation example
// Common Searches: Workbook.CalculateFormula default Aspose.Cells | how to recalculate all formulas in a .NET workbook | Aspose.Cells evaluate formulas without options | C# calculate dependent Excel cells using Aspose | recalculate workbook after changing cell values
// Developer Intent: The developer wants to trigger a full workbook recalculation using Aspose.Cells' standard calculation behavior, without supplying custom options.
// Use Cases: Refresh every formula after programmatically updating cell data before exporting the file. | Generate reports where all calculated values must be resolved automatically. | Validate that newly added or modified formulas produce correct results in a fresh workbook.
// AI Prompts: Show how to recalculate a specific range of cells with Aspose.Cells. | Explain how to customize calculation options for Workbook.CalculateFormula in C#. | Provide code for handling formula calculation errors in Aspose.Cells.

using System;
using Aspose.Cells;

// This C# example creates a workbook, writes a value to A1, adds dependent formulas to B1 and C1, then calls Workbook.CalculateFormula() with no parameters to evaluate every formula using Aspose.Cells' built‑in default calculation engine. The results are printed and the workbook is saved.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet and its cells
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate some data and formulas
        cells["A1"].PutValue(5);               // Simple value
        cells["B1"].Formula = "=A1*2";         // Depends on A1
        cells["C1"].Formula = "=B1+10";        // Depends on B1

        // Calculate all formulas using the default calculation settings
        workbook.CalculateFormula();

        // Display the calculated results
        Console.WriteLine("A1 value: " + cells["A1"].Value);
        Console.WriteLine("B1 value: " + cells["B1"].Value);
        Console.WriteLine("C1 value: " + cells["C1"].Value);

        // Save the workbook (optional)
        workbook.Save("CalculatedWorkbook.xlsx");
    }
}
