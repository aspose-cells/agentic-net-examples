// Title: Round formula results to the displayed number format by enabling PrecisionAsDisplayed in Aspose.Cells for .NET
// AI Prompts: Create a new Workbook, set Workbook.Settings.FormulaSettings.PrecisionAsDisplayed = true, format cell A1 with Number = 2, assign a formula to B1 that references A1, calculate the workbook, and print the displayed values. | Write C# code that saves the workbook after turning on PrecisionAsDisplayed and applying a two‑decimal number format, then verify the saved file retains the rounded values. | Demonstrate how to compare the raw calculated value of B1 with the displayed value of A1 when PrecisionAsDisplayed is enabled, ensuring they match the cell's format.
// Common Searches: Aspose.Cells how to enable rounding to the cell's number format in .NET | C# example for PrecisionAsDisplayed property in Aspose.Cells | Round calculated formula values to two decimal places using Aspose.Cells | Persist displayed precision setting when saving an Excel file with Aspose.Cells | Difference between PrecisionAsDisplayed and default calculation precision in Aspose.Cells
// Tags: Aspose.Cells PrecisionAsDisplayed rounding | formula calculation with displayed precision | apply number format before workbook.CalculateFormula | C# enable displayed precision in Excel workbook | persist rounding setting when saving Aspose.Cells file

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to turn on PrecisionAsDisplayed to round formula results to the cell's displayed number format, apply a two‑decimal format, calculate, and save the workbook while preserving the rounding behavior.
    public class PrecisionAsDisplayedDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Enable rounding to displayed format
                workbook.Settings.FormulaSettings.PrecisionAsDisplayed = true;

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Set a value with many decimal places in A1
                cells["A1"].PutValue(1.23456);

                // Apply a number format to display only 2 decimal places
                Style style = cells["A1"].GetStyle();
                style.Number = 2; // Format: 0.00
                cells["A1"].SetStyle(style);

                // Set a formula in B1 that references A1
                cells["B1"].Formula = "=A1";

                // Calculate formulas (PrecisionAsDisplayed will round the result)
                workbook.CalculateFormula();

                // Output the displayed value of A1 and the calculated value of B1
                Console.WriteLine("A1 Display Value: " + cells["A1"].StringValue); // Expected: 1.23
                Console.WriteLine("B1 Calculated Value: " + cells["B1"].Value);   // Expected: 1.23

                // Save the workbook to demonstrate the setting persists
                workbook.Save("PrecisionAsDisplayedDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            PrecisionAsDisplayedDemo.Run();
        }
    }
}
