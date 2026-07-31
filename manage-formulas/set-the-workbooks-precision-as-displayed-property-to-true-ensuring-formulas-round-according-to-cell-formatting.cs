// Title: C# Aspose.Cells: Set Precision As Displayed to Round Formula Results
// Description: Shows how to enable the PrecisionAsDisplayed option in a .NET workbook, apply a number format, insert a high‑precision value, reference it with a formula, recalculate, and save so that both cells display the rounded value.
// Keywords: Aspose.Cells PrecisionAsDisplayed | C# Excel rounding formulas | displayed precision setting | formula calculation rounding | Aspose.Cells workbook settings | Excel number format rounding | C# .NET Aspose.Cells example | GitHub Aspose.Cells PrecisionAsDisplayed
// Common Searches: Aspose.Cells PrecisionAsDisplayed C# | round Excel formulas with Aspose.Cells | set displayed precision in .NET workbook | enable formula rounding in Aspose.Cells | Aspose.Cells calculate with displayed precision
// Developer Intent: Activate the workbook’s displayed‑precision mode so that all formula evaluations honor the cell’s visible number format.
// Use Cases: Prepare financial statements where every calculated cell must match the two‑decimal display required by auditors. | Export data to downstream systems that only accept rounded values, ensuring consistency by turning on PrecisionAsDisplayed before saving. | Apply custom number formats across a worksheet and enforce uniform rounding for all dependent formulas.
// AI Prompts: Demonstrate toggling PrecisionAsDisplayed on an existing Aspose.Cells workbook and re‑calculating the formulas. | Provide a C# snippet that sets a custom number format, enables displayed precision, and verifies that the formula result respects the rounding. | Explain how PrecisionAsDisplayed affects iterative calculations and how to disable it after processing.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to enable the PrecisionAsDisplayed option in a .NET workbook, apply a number format, insert a high‑precision value, reference it with a formula, recalculate, and save so that both cells display the rounded value.
    public class SetPrecisionAsDisplayedDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Enable PrecisionAsDisplayed so calculations use the displayed precision
            workbook.Settings.FormulaSettings.PrecisionAsDisplayed = true;

            // Access the first worksheet and its cells
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Put a value with many decimal places into A1
            cells["A1"].PutValue(1.23456);

            // Set the number format of A1 to show 2 decimal places (0.00)
            Style style = cells["A1"].GetStyle();
            style.Number = 2;
            cells["A1"].SetStyle(style);

            // Set a formula in B1 that references A1
            cells["B1"].Formula = "=A1";

            // Calculate formulas; with PrecisionAsDisplayed enabled, B1 will use the rounded value
            workbook.CalculateFormula();

            // Display results
            Console.WriteLine("A1 displayed value: " + cells["A1"].StringValue); // Expected: 1.23
            Console.WriteLine("B1 calculated value: " + cells["B1"].Value);      // Expected: 1.23

            // Save the workbook
            workbook.Save("PrecisionAsDisplayedDemo.xlsx");
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                SetPrecisionAsDisplayedDemo.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
