// Title: Show Formulas Instead of Values in Aspose.Cells Worksheet (C#/.NET)
// Description: Shows how to enable Worksheet.ShowFormulas in Aspose.Cells for .NET so cells display their formula strings rather than evaluated results. The example creates a workbook, writes a formula, optionally calculates it, toggles formula view, and saves the file.
// Keywords: Aspose.Cells | Worksheet.ShowFormulas | display formulas | C# .NET | debug Excel formulas | show formulas in workbook | Excel formula view | toggle formula display | programmatic Excel debugging | Aspose.Cells example
// Common Searches: Aspose.Cells show formulas | Worksheet.ShowFormulas property C# | display formula text instead of value Aspose.Cells | debug Excel formulas with Aspose.Cells | how to view formulas in generated Excel file using Aspose.Cells | C# enable formula view in workbook
// Developer Intent: Activate formula view for a worksheet so cells reveal their formula strings rather than computed values, facilitating debugging and verification.
// Use Cases: Inspect formulas while building spreadsheet automation scripts. | Generate an Excel file where formulas are visible for reviewers or auditors. | Toggle between formula view and normal value view before distributing the workbook. | Validate that a formula calculates correctly by comparing its text with the result.
// AI Prompts: Generate C# code that creates a workbook, adds a formula to a cell, enables Worksheet.ShowFormulas, and saves the file. | Explain how Worksheet.ShowFormulas changes cell rendering and how to revert to normal value display. | Provide a snippet that switches ShowFormulas on for debugging and off for final output in the same workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsShowFormulasDemo
{
    // Shows how to enable Worksheet.ShowFormulas in Aspose.Cells for .NET so cells display their formula strings rather than evaluated results. The example creates a workbook, writes a formula, optionally calculates it, toggles formula view, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Set a formula in cell A1
            cells["A1"].Formula = "=1+2+3";

            // Calculate formulas so that the workbook has a valid result (optional)
            workbook.CalculateFormula();

            // Enable formula display for debugging purposes
            worksheet.ShowFormulas = true;

            // The cell now shows the formula text instead of the calculated value
            Console.WriteLine("Displayed in A1: " + cells["A1"].StringValue);

            // Save the workbook (optional, to view in Excel)
            workbook.Save("ShowFormulasDemo.xlsx");
        }
    }
}
