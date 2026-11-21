// Title: How to set manual calculation mode, modify cells, and force a full recalculation with Aspose.Cells in C#
// Common Searches: Aspose.Cells C# set workbook calculation mode to manual | Update cell values and recalculate formulas using Aspose.Cells .NET | Trigger full workbook recalculation after data changes with Aspose.Cells | Calculate all formulas on demand in an Excel file via Aspose.Cells C# | Manual CalcMode example for Aspose.Cells in a .NET application
// Tags: Aspose.Cells manual calculation mode C# | Workbook.CalculateFormula method usage | Updating cell values with Aspose.Cells API | Forcing full formula recalculation in .xlsx files | Setting CalcMode to Manual in Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates creating or loading an Aspose.Cells Workbook, optionally switching its calculation mode to Manual, updating cell A1 with a numeric value and cell B1 with a dependent formula, invoking CalculateFormula to perform a full recalculation, and saving the workbook to an output .xlsx file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // new workbook
            // If you need to load an existing file, ensure it exists:
            // string inputPath = "input.xlsx";
            // if (File.Exists(inputPath))
            //     workbook = new Workbook(inputPath);
            // else
            //     Console.WriteLine($"Input file '{inputPath}' not found. Using a new workbook.");

            // Set calculation mode to Manual if the property exists in the used version
            // workbook.Settings.CalcMode = CalcMode.Manual;

            // Modify data in the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(10);          // set a value
            sheet.Cells["B1"].Formula = "=A1*2";    // set a formula dependent on A1

            // Trigger a full recalculation
            workbook.CalculateFormula();

            // Save the workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
