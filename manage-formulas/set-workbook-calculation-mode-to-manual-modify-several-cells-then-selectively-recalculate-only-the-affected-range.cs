// Title: How to set manual calculation mode and recalculate a specific cell range with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set workbook calculation mode to manual | Recalculate only a range of formulas after updating cells using Aspose.Cells | Partial formula evaluation in Aspose.Cells .NET after cell changes | Use CalculateFormula method for a specific range in Aspose.Cells C# | Load or create a workbook then save with manual calculation mode in Aspose.Cells
// Tags: set manual calculation mode Aspose.Cells | partial formula recalculation Aspose.Cells | CalculateFormula range Aspose.Cells | Workbook.Settings.CalculationMode C# | update cell values Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook or creates a new one, optionally switches the calculation mode to Manual, updates cells A1 and B2, assigns a formula to C3, recalculates only the formulas in the range C3:D5 using CalculateFormula, and saves the result to output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Define input file (optional). If it exists, load it; otherwise create a new workbook.
            string inputPath = "input.xlsx";
            Workbook workbook;

            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
            }

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Example: set calculation mode to Manual if needed
            // workbook.Settings.CalcMode = CalcMode.Manual;                     // older versions
            // workbook.Settings.CalculationMode = CalculationMode.Manual;     // newer versions

            // Modify several cells
            sheet.Cells["A1"].PutValue(100);
            sheet.Cells["B2"].PutValue(200);
            sheet.Cells["C3"].Formula = "=A1+B2";

            // Recalculate only the specified range (C3:D5)
            sheet.CalculateFormula("C3:D5");

            // Ensure the output directory exists (if a directory part is present)
            string outputPath = "output.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath) ?? string.Empty;
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
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
