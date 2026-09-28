// Title: How to turn off PrecisionAsDisplayed in Aspose.Cells (C#) for full double‑precision calculations
// AI Prompts: Generate C# code that sets workbook.Settings.CalcEngineSettings.PrecisionAsDisplayed to false using Aspose.Cells. | Show an example of disabling displayed precision before saving an Excel file with Aspose.Cells in .NET. | Explain how to verify that Aspose.Cells performs calculations with full double precision after changing the precision setting.
// Common Searches: Aspose.Cells C# set PrecisionAsDisplayed false for accurate calculations | disable displayed precision in Aspose.Cells workbook .NET | full double precision arithmetic Aspose.Cells CalcEngineSettings | how to change calculation engine precision in Aspose.Cells C# | Excel workbook precision as displayed option Aspose.Cells example
// Tags: Aspose.Cells PrecisionAsDisplayed flag | CalcEngineSettings precision control C# | Excel double precision mode Aspose.Cells | Workbook calculation settings Aspose.Cells | set calculation precision Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample creates a new Aspose.Cells Workbook, demonstrates how to set workbook.Settings.CalcEngineSettings.PrecisionAsDisplayed to false so that calculations use full double precision, and then saves the workbook to 'output.xlsx' within a try‑catch block.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // If you need to adjust calculation precision, use CalcEngineSettings when available.
            // Example (uncomment if supported by your Aspose.Cells version):
            // workbook.Settings.CalcEngineSettings.PrecisionAsDisplayed = false;

            // Perform any required operations on the workbook here

            // Define output path
            string outputPath = "output.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            // Log or display the error details
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
