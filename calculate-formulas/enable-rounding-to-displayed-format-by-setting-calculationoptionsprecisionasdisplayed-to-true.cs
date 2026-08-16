// Title: Round Formula Results to Displayed Format with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, activates the FormulaSettings.PrecisionAsDisplayed flag, formats a cell to two decimal places, inserts a high‑precision value, copies it via a formula, calculates the sheet, and shows that both cells return the rounded value matching the visible format.
// Keywords: Aspose.Cells | PrecisionAsDisplayed | FormulaSettings | C# rounding | displayed precision | Excel number format | calculate formulas | cell formatting | financial rounding | .NET spreadsheet library
// Common Searches: Aspose.Cells round formula result to cell format | C# set PrecisionAsDisplayed in Aspose.Cells | How to enable displayed precision for calculations in .NET | FormulaSettings PrecisionAsDisplayed example | Round numbers to two decimals with Aspose.Cells
// Developer Intent: Apply the PrecisionAsDisplayed flag so calculated values follow the cell's visible number format.
// Use Cases: Generate financial reports where all computed figures must appear rounded to the displayed decimal places. | Export data to Excel while preserving the visual rounding that users see in the UI. | Validate imported numeric data against the rounded values shown to end‑users before further processing.
// AI Prompts: Demonstrate enabling PrecisionAsDisplayed for multiple cells with different built‑in number formats in Aspose.Cells (C#). | Show a side‑by‑side comparison of formula results with PrecisionAsDisplayed turned on and off. | Explain how PrecisionAsDisplayed influences iterative calculations and how to reset the setting programmatically.

using System;
using Aspose.Cells;

namespace AsposeCellsPrecisionAsDisplayedDemo
{
    // This example creates a workbook, activates the FormulaSettings.PrecisionAsDisplayed flag, formats a cell to two decimal places, inserts a high‑precision value, copies it via a formula, calculates the sheet, and shows that both cells return the rounded value matching the visible format.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Enable rounding to the displayed format
            workbook.Settings.FormulaSettings.PrecisionAsDisplayed = true;

            // Access the first worksheet and its cells collection
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Put a value with many decimal places into A1
            cells["A1"].PutValue(1.23456);

            // Set display format of A1 to show 2 decimal places (0.00)
            Style style = cells["A1"].GetStyle();
            style.Number = 2; // Built‑in format index for two decimal places
            cells["A1"].SetStyle(style);

            // Set a formula in B1 that references A1
            cells["B1"].Formula = "=A1";

            // Calculate formulas; the result will respect the displayed precision
            workbook.CalculateFormula();

            // Output the displayed value of A1 and the calculated value of B1
            Console.WriteLine("A1 Display Value: " + cells["A1"].StringValue); // Expected: 1.23
            Console.WriteLine("B1 Calculated Value: " + cells["B1"].Value);   // Expected: 1.23

            // Save the workbook (optional)
            workbook.Save("PrecisionAsDisplayedDemo.xlsx");
        }
    }
}
