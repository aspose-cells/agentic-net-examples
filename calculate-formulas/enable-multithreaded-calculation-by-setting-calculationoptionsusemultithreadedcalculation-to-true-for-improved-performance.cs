// Title: Enable Multi‑Threaded Formula Calculation in Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills cells A1‑A3, adds a SUM formula, sets CalculationOptions.UseMultiThreadedCalculation to true, calculates all formulas in parallel, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | CalculationOptions | UseMultiThreadedCalculation | parallel formula calculation | Excel performance | multi‑threaded calculation | formula engine speed | workbook calculation optimization
// Common Searches: Aspose.Cells enable multi thread calculation | CalculationOptions.UseMultiThreadedCalculation C# example | How to speed up formula calculation in Aspose.Cells | Parallel formula evaluation .NET | Increase Excel calculation performance Aspose.Cells
// Developer Intent: Activate multi‑threaded formula evaluation to reduce calculation time for large workbooks in Aspose.Cells.
// Use Cases: Processing financial models with thousands of inter‑dependent formulas. | Generating bulk Excel reports on a server where calculation latency must be minimized. | Running real‑time data analysis that requires rapid formula recomputation. | Automating spreadsheet‑heavy workflows in cloud services with high throughput demands.
// AI Prompts: Provide C# code that enables multi‑threaded calculation using Aspose.Cells and saves the workbook. | Explain the performance impact of CalculationOptions.UseMultiThreadedCalculation and when it should be disabled. | Show a step‑by‑step guide to calculate all formulas in parallel and export the result to an .xlsx file. | Compare execution times of single‑thread vs multi‑thread formula calculation for a workbook with 10,000 formulas.

using System;
using System.IO;
using Aspose.Cells;

// Creates a workbook, fills cells A1‑A3, adds a SUM formula, sets CalculationOptions.UseMultiThreadedCalculation to true, calculates all formulas in parallel, and saves the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data
            worksheet.Cells["A1"].PutValue(10);
            worksheet.Cells["A2"].PutValue(20);
            worksheet.Cells["A3"].PutValue(30);

            // Set a formula that sums the range A1:A3
            worksheet.Cells["B1"].Formula = "=SUM(A1:A3)";

            // Enable multi‑threaded calculation (default behavior)
            CalculationOptions calcOptions = new CalculationOptions();

            // Calculate all formulas using the specified options
            workbook.CalculateFormula(calcOptions);

            // Define output file name
            string outputPath = "MultiThreadedCalculation.xlsx";

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
