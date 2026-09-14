// Title: Disable EnableThreadedCalculation in Aspose.Cells C# to lower calculation overhead for simple workbooks
// AI Prompts: Generate C# code that uses reflection to set Workbook.Settings.EnableThreadedCalculation to false before saving the file. | Create a minimal Aspose.Cells workbook with a SUM formula, turn off threaded calculation, calculate the formula, and save it as an .xlsx file. | Describe the stability benefits of disabling threaded calculation for small workbooks and outline a version‑compatible approach.
// Common Searches: C# Aspose.Cells how to turn off threaded calculation for small workbooks | set EnableThreadedCalculation false using reflection Aspose.Cells | reduce formula calculation overhead Aspose.Cells workbook settings | Aspose.Cells stability issue with threaded calculation in .NET | disable multithreaded calculation Aspose.Cells example
// Tags: disable EnableThreadedCalculation Aspose.Cells | Aspose.Cells workbook settings performance | C# reflection modify Aspose.Cells property | threaded calculation overhead Excel generation | stability improvement simple workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// // Demonstrates creating a new Workbook, using reflection to set EnableThreadedCalculation to false for compatibility, adding sample data with a SUM formula, calculating the formula, and saving the workbook.
class ReduceThreadedCalculationDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook instance
            Workbook workbook = new Workbook();

            // Attempt to disable threaded calculation via reflection (property may not exist in some versions)
            try
            {
                var settingsType = workbook.Settings.GetType();
                var prop = settingsType.GetProperty("EnableThreadedCalculation");
                if (prop != null && prop.CanWrite)
                {
                    prop.SetValue(workbook.Settings, false);
                }
            }
            catch (Exception ex)
            {
                // If reflection fails, continue without disabling threaded calculation
                Console.WriteLine($"Unable to set EnableThreadedCalculation: {ex.Message}");
            }

            // Add sample data and a formula
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].Formula = "=SUM(A1:A2)";

            // Calculate formulas (optional)
            workbook.CalculateFormula();

            // Define output file path
            string outputPath = "SimpleWorkbook.xlsx";

            // Ensure the output directory exists
            string directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
