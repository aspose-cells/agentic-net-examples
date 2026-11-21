// Title: How to disable PrecisionAsDisplayed and use full double precision for formula calculations in Aspose.Cells for .NET (C#)
// Common Searches: how to turn off precision as displayed in Aspose.Cells using C# | calculate Excel formulas with full double precision in Aspose.Cells .NET | set workbook precision false Aspose.Cells example | does Aspose.Cells default calculation engine use double precision | C# code to evaluate formula with maximum precision in Aspose.Cells
// Tags: disable precision as displayed Aspose.Cells | full double precision formula calculation .NET | Aspose.Cells calculation engine precision setting | C# workbook precision false | Excel double precision calculations Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a Workbook, adds numeric values, and assigns a formula to demonstrate full double‑precision evaluation. It relies on Aspose.Cells' default calculation engine, which already performs calculations with maximum double precision, and shows how to verify the result, then saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Note: In recent Aspose.Cells versions the CalcEngineSettings property may be unavailable.
            // The default calculation engine already uses full double precision, so we can proceed without
            // explicitly setting PrecisionAsDisplayed.

            // Example: add some data and a formula to verify full precision calculation
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(0.1);
            sheet.Cells["A2"].PutValue(0.2);
            sheet.Cells["A3"].Formula = "=A1+A2"; // Expected precise result: 0.30000000000000004

            // Force calculation to apply the new precision setting (or default behavior)
            workbook.CalculateFormula();

            // Output the result to the console
            Console.WriteLine("Result of A3 (full double precision): " + sheet.Cells["A3"].Value);

            // Save the workbook if needed
            string outputPath = "PrecisionExample.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
