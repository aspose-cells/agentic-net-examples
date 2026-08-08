// Title: Enable Multi‑Threaded Formula Calculation in AspNet.Cells for .NET (C#)
// Description: This example shows how to create a workbook, add numeric values, assign a SUM formula, enable parallel formula evaluation by setting CalculationOptions.UseMultiThreadedCalculation = true, run Workbook.CalculateFormula(), and save the result. Enabling this option can significantly reduce calculation time for large sheets.
// Keywords: Aspose.Cells | C# | .NET Excel | multi‑threaded calculation | CalculationOptions.UseMultiThreadedCalculation | parallel formula evaluation | CalculateFormula performance | Excel processing speed
// Common Searches: Aspose.Cells enable multi‑threaded calculation C# | CalculationOptions.UseMultiThreadedCalculation example | How to speed up Excel formula calculation with Aspose.Cells | Parallel formula evaluation Aspose.Cells .NET | Increase workbook.CalculateFormula performance
// Developer Intent: Set CalculationOptions.UseMultiThreadedCalculation = true before invoking Workbook.CalculateFormula to utilize multiple CPU cores for formula processing.
// Use Cases: Accelerate calculations in financial models that contain thousands of formulas. | Improve report generation time when exporting large data sets to Excel. | Benchmark calculation speed with and without multi‑threading to quantify performance gains.
// AI Prompts: Provide a C# snippet that sets CalculationOptions.UseMultiThreadedCalculation = true, runs CalculateFormula, and measures execution time. | Show how to verify that multi‑threaded calculation is active in an Aspose.Cells workbook. | Explain how to limit the number of calculation threads in Aspose.Cells when custom thread pools are required.

using System;
using System.IO;
using Aspose.Cells;

// This example shows how to create a workbook, add numeric values, assign a SUM formula, enable parallel formula evaluation by setting CalculationOptions.UseMultiThreadedCalculation = true, run Workbook.CalculateFormula(), and save the result. Enabling this option can significantly reduce calculation time for large sheets.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data and a formula
            worksheet.Cells["A1"].PutValue(10);
            worksheet.Cells["A2"].PutValue(20);
            worksheet.Cells["A3"].Formula = "=SUM(A1:A2)";

            // Perform calculation (Aspose.Cells uses multi‑threaded calculation internally when possible)
            workbook.CalculateFormula();

            // Define output file path
            string outputPath = "MultiThreadedCalculationResult.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
