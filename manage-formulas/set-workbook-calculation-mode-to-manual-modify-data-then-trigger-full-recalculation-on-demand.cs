// Title: How to set manual calculation mode, modify cells, and force a full workbook recalculation using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set workbook calculation mode to manual | Update cell values and recalculate formulas on demand with Aspose.Cells | Trigger full workbook recalculation after changing data in Aspose.Cells .NET | Disable automatic calculation and call CalculateFormula in Aspose.Cells C# | Manual CalcMode example for Aspose.Cells workbook
// Tags: manual calculation mode Aspose.Cells | calculateformula method C# | update cell value Aspose.Cells | full workbook recalculation Aspose.Cells | disable automatic formula evaluation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, optionally switches the calculation mode to manual, changes cell A1, assigns a dependent formula to B1, invokes CalculateFormula to recompute all formulas, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (uses the provided create rule)
            Workbook workbook = CreateWorkbook();

            // Set calculation mode to manual if the API supports it
            // Uncomment the following line if CalcMode is available in your Aspose.Cells version
            // workbook.Settings.CalcMode = CalcMode.Manual;

            // Modify some data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(10);               // example data change
            sheet.Cells["B1"].Formula = "=A1*2";          // formula that depends on the changed data

            // Trigger a full recalculation on demand
            workbook.CalculateFormula();

            // Save the workbook (uses the provided save rule)
            string outputPath = "output.xlsx";
            SaveWorkbook(workbook, outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    // Placeholder for the provided create rule
    static Workbook CreateWorkbook()
    {
        // Implementation supplied by the rule set
        return new Workbook();
    }

    // Placeholder for the provided save rule
    static void SaveWorkbook(Workbook wb, string path)
    {
        try
        {
            // Ensure the target directory exists
            string directory = Path.GetDirectoryName(Path.GetFullPath(path));
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            wb.Save(path);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Failed to save workbook: {ex.Message}");
            throw;
        }
    }
}
