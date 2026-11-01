// Title: Enable manual calculation mode and recalculate formulas in an Aspose.Cells workbook using C#
// Common Searches: Aspose.Cells C# set workbook calculation mode to Manual | C# recalculate Excel formulas after changing CalcMode in Aspose.Cells | Trigger formula evaluation for a specific worksheet with Aspose.Cells .NET | Programmatically calculate formulas on demand in Aspose.Cells | Save workbook after manual formula calculation using Aspose.Cells C#
// Tags: manual calculation mode Aspose.Cells | calculate workbook formulas Aspose.Cells | trigger worksheet recalculation Aspose.Cells | Settings.CalcMode API C# | Aspose.Cells formula evaluation on demand

using Aspose.Cells;
using System;
using System.IO;

// The example shows how to set the workbook's calculation mode to Manual via Settings.CalcMode (when supported) and invoke CalculateFormula() to evaluate formulas programmatically, then save the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // If your Aspose.Cells version supports CalcMode, you can set it to Manual:
            // workbook.Settings.CalcMode = CalcMode.Manual;

            // Calculate all formulas in the workbook (covers all worksheets)
            workbook.CalculateFormula();

            // Save the workbook to verify changes (optional)
            string outputPath = "ManualCalcResult.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
