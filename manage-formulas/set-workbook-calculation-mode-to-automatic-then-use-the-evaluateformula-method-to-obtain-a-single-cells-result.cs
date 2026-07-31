// Title: Aspose.Cells C# – Set Workbook to Automatic Calculation and Evaluate a Formula with CalculateFormula
// Description: Demonstrates how to enable automatic calculation mode in an Aspose.Cells workbook and use the CalculateFormula method to evaluate a formula (e.g., =SUM(A1:A3)) directly, without writing it to a worksheet cell, in C# .NET.
// Keywords: Aspose.Cells automatic calculation mode | C# CalculateFormula example | evaluate formula without cell Aspose.Cells | Aspose.Cells workbook settings | Sum formula evaluation .NET | Aspose.Cells CalcModeType.Automatic | programmatic formula evaluation C#
// Common Searches: Aspose.Cells set calculation mode to automatic C# | How to evaluate a formula without placing it in a cell using Aspose.Cells | CalculateFormula method example Aspose.Cells .NET | Get result of SUM formula programmatically Aspose.Cells | Enable automatic workbook recalculation Aspose.Cells
// Developer Intent: Enable automatic calculation for a workbook and retrieve the result of a formula programmatically using Aspose.Cells in C#.
// Use Cases: Generate summary totals on the fly for reports without persisting formulas. | Validate data by evaluating custom expressions after populating cells in memory. | Perform quick, in‑memory calculations before deciding to save or export the workbook.
// AI Prompts: Provide C# code that sets Aspose.Cells workbook calculation mode to Automatic and evaluates a formula with CalculateFormula. | Explain how to use CalculateFormula to obtain a numeric result from a formula without inserting it into a worksheet cell. | Show how to handle different return types (numeric, string, error) from sheet.CalculateFormula in Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to enable automatic calculation mode in an Aspose.Cells workbook and use the CalculateFormula method to evaluate a formula (e.g., =SUM(A1:A3)) directly, without writing it to a worksheet cell, in C# .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Set the workbook calculation mode to Automatic
        wb.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

        // Get the first worksheet and its cells collection
        Worksheet sheet = wb.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate some sample data
        cells["A1"].PutValue(5);
        cells["A2"].PutValue(10);
        cells["A3"].PutValue(15);

        // Evaluate a formula directly without placing it in a cell
        object result = sheet.CalculateFormula("=SUM(A1:A3)");

        // Output the evaluated result
        Console.WriteLine("Result of =SUM(A1:A3): " + result);

        // Save the workbook (optional)
        wb.Save("ResultWorkbook.xlsx");
    }
}
