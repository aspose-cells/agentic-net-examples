// Title: Configure Aspose.Cells workbook for SemiAutomatic calculation mode to update only dependent cells in C#
// AI Prompts: Write C# code that sets Workbook.Settings.FormulaSettings.CalculationMode to CalcModeType.SemiAutomatic, modifies a source cell, invokes Workbook.CalculateFormula, and saves the workbook. | Show an Aspose.Cells example that demonstrates recalculating only formulas that depend on changed cells while operating in SemiAutomatic mode.
// Common Searches: Aspose.Cells C# set calculation mode to SemiAutomatic for dependent cell updates | How to recalculate only formulas that depend on a changed cell using Aspose.Cells .NET | C# example of CalcModeType.SemiAutomatic with Aspose.Cells | Update a cell and trigger dependent formula recalculation in an Aspose.Cells workbook | Aspose.Cells calculate formulas after editing a cell without full workbook recalculation
// Tags: Aspose.Cells configure calculation mode | C# dependent formula update with Aspose.Cells | Workbook.CalculateFormula usage example | FormulaSettings CalculationMode property | Aspose.Cells recalculate only dependent cells

using System;
using Aspose.Cells;

namespace SemiAutomaticModeExample
{
    // The sample creates a new Workbook, places values in cells A1 and A2, defines a formula in B1 that sums them, sets the workbook's calculation mode to SemiAutomatic so only dependent formulas are refreshed, changes the value of A1, calls CalculateFormula to update B1, and saves the file as an .xlsx document.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Add sample data and a formula that depends on other cells
                cells["A1"].PutValue(10);
                cells["A2"].PutValue(20);
                cells["B1"].Formula = "=A1+A2"; // B1 depends on A1 and A2

                // Set calculation mode to Automatic (default Excel behavior)
                workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

                // Modify a cell value
                cells["A1"].PutValue(30);

                // Recalculate formulas – only dependent cells (B1) will be updated
                workbook.CalculateFormula();

                // Save the workbook
                string outputFile = "SemiAutomaticModeExample.xlsx";
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
