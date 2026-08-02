// Title: Aspose.Cells .NET: Disable Automatic Formula Recalculation for Faster Large Data Import
// Description: Shows how to set a workbook to manual calculation, turn off CalculateOnOpen/Save, import a 100,000‑row DataTable with ImportData, and save without triggering formula evaluation.
// Keywords: Aspose.Cells manual calculation | disable formula recalculation | bulk data import performance | ImportData large dataset | CalcModeType.Manual | CalculateOnOpen false | CalculateOnSave false | .NET spreadsheet library | optimize workbook save speed
// Common Searches: how to turn off formula calculation in Aspose.Cells .NET | increase import speed Aspose.Cells large DataTable | disable calculate on open Aspose.Cells | manual formula mode bulk import | optimize Aspose.Cells performance for big worksheets
// Developer Intent: Prevent automatic formula evaluation during massive data loads to improve import throughput.
// Use Cases: Import a 100k‑row DataTable into a worksheet while keeping formulas unevaluated, then save the file. | Load an existing workbook, switch to manual mode, modify cells, and invoke wb.CalculateFormula() only when needed. | Generate a numeric report, disable CalculateOnOpen/Save to reduce processing time, and re‑enable automatic calculation after final edits.
// AI Prompts: Provide C# code that sets Aspose.Cells to manual calculation, imports a large DataTable, and later runs CalculateFormula. | Explain how to toggle CalculateOnOpen and CalculateOnSave flags for performance optimization during bulk imports. | Show the steps to switch back to automatic formula evaluation after importing a massive data set with Aspose.Cells.

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Shows how to set a workbook to manual calculation, turn off CalculateOnOpen/Save, import a 100,000‑row DataTable with ImportData, and save without triggering formula evaluation.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook wb = new Workbook();

            // Disable automatic formula calculation to speed up large data imports
            wb.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;
            wb.Settings.FormulaSettings.CalculateOnOpen = false;
            wb.Settings.FormulaSettings.CalculateOnSave = false;

            // Access the first worksheet
            Worksheet sheet = wb.Worksheets[0];
            Cells cells = sheet.Cells;

            // Simulate a large data set (e.g., 100,000 rows × 5 columns)
            int rows = 100_000;
            int cols = 5;

            // Build a DataTable with the simulated data
            DataTable dt = new DataTable();
            for (int c = 0; c < cols; c++)
            {
                dt.Columns.Add("Col" + c, typeof(double));
            }

            for (int r = 0; r < rows; r++)
            {
                DataRow dr = dt.NewRow();
                for (int c = 0; c < cols; c++)
                {
                    dr[c] = r * cols + c + 1; // simple numeric value
                }
                dt.Rows.Add(dr);
            }

            // Import the DataTable into the worksheet starting at cell A1
            // (lifecycle: import)
            ImportTableOptions importOptions = new ImportTableOptions(); // default options
            cells.ImportData(dt, 0, 0, importOptions);

            // If you need to calculate formulas later, call CalculateFormula explicitly
            // wb.CalculateFormula();

            // Save the workbook (lifecycle: save)
            wb.Save("LargeDataManualCalc.xlsx", SaveFormat.Xlsx);
        }
    }
}
