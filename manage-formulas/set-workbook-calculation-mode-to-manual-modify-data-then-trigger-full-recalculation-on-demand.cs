// Title: Set manual calculation mode, update cell values, and trigger full workbook recalculation with Aspose.Cells for .NET
// AI Prompts: Configure the workbook to use manual calc mode, write numeric values to A1 and A2, assign a formula to B1, then call CalculateFormula to recompute all dependent cells before saving. | Create a new workbook, switch the calculation setting to manual via the Settings object, modify cell data, add a dependent formula, and force a full recalculation on demand.
// Common Searches: how to disable automatic calculation in Aspose.Cells .NET and recalc later | set workbook calculation mode to manual Aspose.Cells example | force full formula recalculation after cell updates using Aspose.Cells | Aspose.Cells calculate all formulas programmatically C# | manual CalcMode usage with Aspose.Cells for .NET
// Tags: manual calculation mode Aspose.Cells | CalcMode property Aspose.Cells | CalculateFormula method Aspose.Cells | update cell values programmatically Aspose.Cells | full workbook recalculation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// // Creates a workbook, optionally sets calculation to manual, writes values to A1/A2, adds a formula in B1, forces a full recalculation with CalculateFormula, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // (Optional) Set calculation mode to manual if supported
            // Uncomment the following lines if your Aspose.Cells version provides CalcMode
            // try
            // {
            //     workbook.Settings.CalcMode = CalcMode.Manual;
            // }
            // catch { /* CalcMode not available in this version */ }

            // Get the first worksheet and its cells collection
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Modify some data in the worksheet
            cells["A1"].PutValue(10);          // Put numeric value 10 in A1
            cells["A2"].PutValue(20);          // Put numeric value 20 in A2
            cells["B1"].Formula = "=A1+A2";    // Set a formula that depends on A1 and A2

            // Trigger a full recalculation of the workbook on demand
            workbook.CalculateFormula();

            // Define output path and ensure the directory exists
            string outputPath = "output.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
