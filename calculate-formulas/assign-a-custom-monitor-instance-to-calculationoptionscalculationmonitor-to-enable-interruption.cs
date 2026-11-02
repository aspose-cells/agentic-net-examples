// Title: How to assign a custom CalculationMonitor to CalculationOptions in Aspose.Cells (C#) for interrupting formula evaluation
// Common Searches: how to cancel workbook.CalculateFormula using CalculationOptions in C# | Aspose.Cells custom CalculationMonitor example for long‑running formulas | set CalculationOptions.CalculationMonitor to stop formula processing in C# | interrupt formula calculation in Aspose.Cells with a user‑defined monitor | using CalculationMonitor to break out of CalculateFormula loop in C#
// Tags: custom calculation monitor Aspose.Cells | CalculationOptions monitor C# | interrupt workbook.CalculateFormula | formula evaluation cancellation Aspose.Cells | implementing CalculationMonitor C#

using System;
using System.IO;
using Aspose.Cells;

// This example demonstrates how to create a custom CalculationMonitor, assign it to CalculationOptions.CalculationMonitor, and use it to interrupt a long‑running workbook.CalculateFormula operation in Aspose.Cells for C#.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Populate cells with data and a formula.
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].Formula = "=SUM(A1:A2)";

            // Calculate formulas.
            try
            {
                workbook.CalculateFormula();
                Console.WriteLine("Formula calculation completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during calculation: {ex.Message}");
            }

            // Define output path and ensure the directory exists.
            string outputPath = "Output.xlsx";
            try
            {
                string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
