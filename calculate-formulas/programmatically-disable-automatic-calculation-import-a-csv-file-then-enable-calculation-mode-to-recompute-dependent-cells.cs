// Title: Disable automatic calculation, import a CSV file, then re‑enable calculation and recalculate formulas with Aspose.Cells for .NET
// AI Prompts: Write C# code that sets Workbook.Settings.FormulaSettings.CalculationMode to Manual, uses Cells.ImportCSV to load a CSV file into the first worksheet, switches the mode back to Automatic, and calls Workbook.CalculateFormula to update all dependent cells. | Provide a step‑by‑step example showing how to pause formula evaluation in an Aspose.Cells workbook, import CSV data, resume automatic calculation, and force a full formula recomputation.
// Common Searches: Aspose.Cells C# import CSV without triggering formula calculation | How to set manual calculation mode before loading data in Aspose.Cells .NET | Recalculate all formulas after importing CSV with Aspose.Cells | Temporarily disable formula evaluation in an Aspose.Cells workbook
// Tags: Aspose.Cells manual calculation mode .NET | Cells.ImportCSV CSV data loading | Aspose.Cells automatic calculation mode | Workbook.CalculateFormula full recalculation | disable formula evaluation Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsCalcModeExample
{
    // The example creates a new workbook, switches the calculation mode to manual, imports CSV data starting at cell A1, restores automatic calculation, forces a full formula recalculation, and saves the result as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet and its cells
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Disable automatic calculation (FormulaSettings.CalculationMode = Manual)
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

            // Path to the CSV file to be imported
            string csvPath = "data.csv";

            // Import CSV data starting at cell A1 (row 0, column 0)
            // Using comma as delimiter and converting numeric data
            cells.ImportCSV(csvPath, ",", true, 0, 0);

            // Re‑enable automatic calculation (or choose another mode as needed)
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

            // Recalculate all formulas after the import
            workbook.CalculateFormula();

            // Save the workbook (lifecycle rule: save)
            workbook.Save("Result.xlsx", SaveFormat.Xlsx);
        }
    }
}
