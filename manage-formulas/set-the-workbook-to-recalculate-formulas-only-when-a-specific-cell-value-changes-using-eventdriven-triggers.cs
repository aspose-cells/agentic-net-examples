// Title: Recalculate workbook formulas only after a specific cell changes using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set calculation mode to manual and recalculate formulas after updating a cell | how to trigger workbook.CalculateFormula only when a particular cell value changes in .NET | event‑driven formula update with Aspose.Cells when cell B2 is modified | disable automatic formula calculation in Aspose.Cells and run manual calculation programmatically | C# example for conditional formula evaluation based on cell change using Aspose.Cells
// Tags: manual calculation mode Aspose.Cells | trigger formula recalculation programmatically | conditional workbook calculation .NET | disable auto formula evaluation Aspose.Cells | calculate formulas after cell update C#

using Aspose.Cells;
using System;
using System.IO;

// The example creates a workbook, optionally switches the calculation mode to Manual, changes cell B2, explicitly calls workbook.CalculateFormula() to recalculate dependent formulas, and saves the file as an XLSX document.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one if needed)
            Workbook workbook = new Workbook();

            // OPTIONAL: Disable automatic calculation if the API version supports it
            // Uncomment the following lines if CalcMode property is available in your Aspose.Cells version
            // workbook.Settings.CalcMode = CalcMode.Manual;

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Modify the trigger cell value programmatically
            sheet.Cells["B2"].PutValue(10);

            // Manually recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
