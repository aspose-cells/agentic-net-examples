// Title: C# – Disable Precision As Displayed in Aspose.Cells for Full Double Precision
// Description: Shows how to set Workbook.Settings.FormulaSettings.PrecisionAsDisplayed to false, run calculations with full double‑precision, and save the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells PrecisionAsDisplayed false | full double precision formulas | disable displayed precision | Aspose.Cells formula settings | C# Excel calculation precision | Aspose.Cells workbook settings | precision as displayed property
// Common Searches: Aspose.Cells set PrecisionAsDisplayed false C# | turn off precision as displayed Aspose.Cells | full double precision calculation Aspose.Cells .NET | how to disable displayed precision in Excel using Aspose | Aspose.Cells formula precision setting example
// Developer Intent: Set the workbook's PrecisionAsDisplayed flag to false so formulas are evaluated with full double‑precision instead of the displayed rounding.
// Use Cases: Create a new workbook, disable PrecisionAsDisplayed, add high‑precision numbers, calculate formulas, and save the file. | Load an existing Excel file, turn off PrecisionAsDisplayed, recalculate all formulas for more accurate financial results, and overwrite the original workbook. | Generate a data‑intensive report where rounding errors must be avoided, ensuring the workbook setting is disabled before any formula evaluation.
// AI Prompts: Write C# code that opens an existing Excel workbook with Aspose.Cells, sets PrecisionAsDisplayed to false, recalculates all formulas, and saves the updated file. | Provide an example that adds high‑precision numeric values to a worksheet, disables PrecisionAsDisplayed, calculates the results, and exports them to CSV using Aspose.Cells. | Explain how the PrecisionAsDisplayed property influences formula evaluation in Aspose.Cells and demonstrate how to toggle it programmatically.

using System;
using Aspose.Cells;

// Shows how to set Workbook.Settings.FormulaSettings.PrecisionAsDisplayed to false, run calculations with full double‑precision, and save the workbook using Aspose.Cells for .NET.
class SetPrecisionAsDisplayed
{
    public static void Main()
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Disable PrecisionAsDisplayed so calculations use full double precision
        workbook.Settings.FormulaSettings.PrecisionAsDisplayed = false;

        // Optional demonstration: add values and a formula
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(1.23456789);
        sheet.Cells["A2"].PutValue(2.34567891);
        sheet.Cells["A3"].Formula = "=A1+A2";

        // Perform calculation with full precision
        workbook.CalculateFormula();

        // Save the workbook
        string outputPath = "PrecisionAsDisplayedFalse.xlsx";
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to {outputPath}");
    }
}
