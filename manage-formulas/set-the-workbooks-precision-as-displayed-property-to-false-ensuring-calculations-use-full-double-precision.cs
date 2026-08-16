// Title: Aspose.Cells C# – Turn Off PrecisionAsDisplayed to Use Full Double Precision
// Description: Shows how to create a Workbook, set Workbook.Settings.FormulaSettings.PrecisionAsDisplayed to false, run calculations with native double‑precision values, and save the result. Ideal for eliminating rounding errors in Excel formulas when using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | PrecisionAsDisplayed | full double precision | C# .NET | formula calculation | disable displayed precision | Excel rounding | Workbook.Settings.FormulaSettings | CalculateFormula | high‑precision calculations
// Common Searches: Aspose.Cells disable PrecisionAsDisplayed | C# set PrecisionAsDisplayed false | full precision formula evaluation Aspose.Cells | prevent rounding in Aspose.Cells calculations | Aspose.Cells high precision mode .NET
// Developer Intent: Turn off the PrecisionAsDisplayed flag so that all formulas are evaluated with native double precision.
// Use Cases: Generate financial models where rounding must be avoided by keeping full numeric accuracy. | Export scientific data to Excel while preserving the exact results of complex calculations. | Create reporting tools that require deterministic results across different locales and Excel versions.
// AI Prompts: Write C# code that disables PrecisionAsDisplayed in Aspose.Cells, runs a sample formula, and prints the computed value. | Explain how setting PrecisionAsDisplayed to false affects Excel formula evaluation and rounding behavior in Aspose.Cells. | Provide a snippet that toggles PrecisionAsDisplayed based on a configuration flag and recalculates the workbook.

using System;
using Aspose.Cells;

// Shows how to create a Workbook, set Workbook.Settings.FormulaSettings.PrecisionAsDisplayed to false, run calculations with native double‑precision values, and save the result. Ideal for eliminating rounding errors in Excel formulas when using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Disable PrecisionAsDisplayed so calculations use full double precision
        workbook.Settings.FormulaSettings.PrecisionAsDisplayed = false;

        // Sample data to demonstrate calculation
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(1.23456);
        sheet.Cells["B1"].Formula = "=A1";

        // Perform calculation
        workbook.CalculateFormula();

        // Save the workbook
        workbook.Save("PrecisionAsDisplayedFalse.xlsx");
    }
}
