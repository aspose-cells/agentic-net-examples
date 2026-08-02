// Title: Aspose.Cells for .NET – Disable Automatic Calculation During Bulk Import and Recalculate Afterwards
// Description: Demonstrates how to set Workbook.Settings.FormulaSettings.CalculationMode to Manual, import 10,000 rows of data with Cells.ImportArray, add SUM formulas, switch back to Automatic mode, force a full recalculation with Workbook.CalculateFormula, and save the workbook. This approach eliminates per‑cell formula evaluation and speeds up large data loads.
// Keywords: Aspose.Cells | C# | .NET | disable automatic calculation | manual calculation mode | bulk data import | ImportArray performance | Workbook.CalculateFormula | Excel formula recalculation | high‑performance Excel export | large dataset import
// Common Searches: how to turn off formula calculation in Aspose.Cells .NET | bulk import Excel data without recalculation Aspose.Cells | set calculation mode to manual then automatic Aspose.Cells | speed up large Excel writes with Aspose.Cells | recalculate all formulas after bulk import Aspose.Cells
// Developer Intent: Temporarily switch off automatic formula calculation while loading massive data, then re‑enable it and run a single full recalculation.
// Use Cases: Load tens of thousands of rows into a worksheet without the overhead of per‑cell formula evaluation. | Add summary or aggregate formulas after a bulk data load and update results in one step. | Create fast‑generated Excel reports by toggling calculation mode before and after data population.
// AI Prompts: Show me C# code to set Aspose.Cells calculation mode to manual, import data with ImportArray, then restore automatic mode and recalculate. | Explain the performance benefits of disabling calculation during large data imports with Aspose.Cells and how to correctly re‑enable it. | Provide an example that adds SUM formulas after a bulk import and forces a full workbook recalculation using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBulkImport
{
    // Demonstrates how to set Workbook.Settings.FormulaSettings.CalculationMode to Manual, import 10,000 rows of data with Cells.ImportArray, add SUM formulas, switch back to Automatic mode, force a full recalculation with Workbook.CalculateFormula, and save the workbook. This approach eliminates per‑cell formula evaluation and speeds up large data loads.
    public class BulkImportExample
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // -----------------------------------------------------------------
            // 1. Disable automatic calculation to speed up bulk data import
            // -----------------------------------------------------------------
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

            // -----------------------------------------------------------------
            // 2. Perform bulk data import (row‑by‑row using string[] overload)
            // -----------------------------------------------------------------
            int rows = 10000;
            int cols = 10;

            for (int r = 0; r < rows; r++)
            {
                string[] rowData = new string[cols];
                for (int c = 0; c < cols; c++)
                {
                    rowData[c] = (r * cols + c + 1).ToString(); // sample values as strings
                }

                // Import the current row starting at column 0
                cells.ImportArray(rowData, r, 0, false);
            }

            // Optionally add some formulas that depend on the imported data
            for (int r = 0; r < rows; r++)
            {
                // Formula: =SUM(A{r+1}:J{r+1})
                string address = $"A{r + 1}";
                string formula = $"=SUM({address}:J{r + 1})";
                cells[r, cols].Formula = formula; // column index 10 (after the 10 data columns)
            }

            // -----------------------------------------------------------------
            // 3. Re‑enable automatic calculation and force a full recalculation
            // -----------------------------------------------------------------
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;
            workbook.CalculateFormula();

            // -----------------------------------------------------------------
            // 4. Save the workbook (lifecycle save)
            // -----------------------------------------------------------------
            string outputPath = "BulkImportResult.xlsx";

            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
    }
}
