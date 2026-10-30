// Title: Enable Fast Formula Calculation in Aspose.Cells for .NET to Accelerate Workbook.CalculateFormula on Large Excel Files
// Common Searches: C# set EnableFastFormulaCalculation before CalculateFormula Aspose.Cells | speed up Aspose.Cells formula evaluation for large workbooks | Aspose.Cells performance tuning using Workbook.Settings.EnableFastFormulaCalculation | how to improve CalculateFormula execution time in .NET with Aspose.Cells | example of fast formula calculation setting in Aspose.Cells C#
// Tags: EnableFastFormulaCalculation Aspose.Cells | Aspose.Cells workbook.CalculateFormula performance | fast formula evaluation .NET Excel | large workbook calculation optimization Aspose.Cells | Workbook.Settings performance flag Aspose.Cells | C# Aspose.Cells formula calculation speed

using System;
using System.IO;
using Aspose.Cells;

// The program loads or creates a workbook, enables the fast formula calculation flag via workbook.Settings.EnableFastFormulaCalculation, executes workbook.CalculateFormula to process all formulas efficiently on large datasets, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            if (File.Exists(inputPath))
            {
                // Load the existing workbook
                workbook = new Workbook(inputPath);
            }
            else
            {
                // Create a new workbook if the input file is missing
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the workbook after calculation
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
