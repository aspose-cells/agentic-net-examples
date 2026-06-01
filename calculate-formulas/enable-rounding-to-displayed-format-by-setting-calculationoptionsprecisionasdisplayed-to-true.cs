using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class PrecisionAsDisplayedDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (in-memory)
                Workbook workbook = new Workbook();

                // Enable rounding to displayed format
                workbook.Settings.FormulaSettings.PrecisionAsDisplayed = true;

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Put a value with many decimal places into A1
                cells["A1"].PutValue(1.23456);

                // Set display format of A1 to show 2 decimal places (0.00)
                Style style = cells["A1"].GetStyle();
                style.Number = 2;
                cells["A1"].SetStyle(style);

                // Set a formula in B1 that references A1
                cells["B1"].Formula = "=A1";

                // Calculate formulas with the PrecisionAsDisplayed option enabled
                workbook.CalculateFormula();

                // Output the displayed value of A1 and the calculated value of B1
                Console.WriteLine("A1 Display Value: " + cells["A1"].StringValue); // Expected: 1.23
                Console.WriteLine("B1 Calculated Value: " + cells["B1"].Value);   // Expected: 1.23

                // Save the workbook; ensure the path is valid
                string outputPath = "PrecisionAsDisplayedDemo.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            PrecisionAsDisplayedDemo.Run();
        }
    }
}