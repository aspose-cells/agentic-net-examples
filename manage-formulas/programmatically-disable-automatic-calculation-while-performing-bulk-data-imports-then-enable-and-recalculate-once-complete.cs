// Title: Disable Automatic Calculation for Bulk Imports and Recalculate with Aspose.Cells for .NET
// Description: Show how to set FormulaSettings.CalculationMode to Manual, import large data sets efficiently, then restore Automatic mode and recalculate all formulas using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | manual calculation mode | bulk data import | disable automatic calculation | calculate formulas | FormulaSettings | performance optimization | large worksheet | Excel automation
// Common Searches: Aspose.Cells turn off calculation | bulk import performance Aspose.Cells | manual calculation mode .NET | recalculate workbook after data load Aspose | disable calculate on save Aspose.Cells
// Developer Intent: Temporarily switch to manual calculation while inserting massive data, then re‑enable automatic mode and evaluate all dependent formulas.
// Use Cases: Import a 10,000‑row numeric array without triggering per‑cell formula evaluation, then add column‑sum formulas and compute them in one step. | Load external data into a workbook with calculation disabled, enable CalculateOnSave before saving to ensure the file contains evaluated results. | Perform multiple worksheet updates in loops, disable automatic calculation for speed, and finally call Workbook.CalculateFormula() to refresh dependent cells.
// AI Prompts: Generate C# code that disables automatic calculation in Aspose.Cells, bulk‑imports a large two‑dimensional array, adds formulas, re‑enables calculation, and recalculates before saving. | Explain how FormulaSettings.CalculationMode and CalculateOnSave work together to improve performance during massive data insertion with Aspose.Cells. | Provide best‑practice tips for optimizing memory and speed when inserting millions of cells using Aspose.Cells, including manual calculation handling.

using System;
using Aspose.Cells;

namespace BulkImportExample
{
    // Show how to set FormulaSettings.CalculationMode to Manual, import large data sets efficiently, then restore Automatic mode and recalculate all formulas using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (lifecycle create)
                Workbook workbook = new Workbook();

                // Access formula settings
                FormulaSettings formulaSettings = workbook.Settings.FormulaSettings;

                // Disable automatic calculation during bulk import
                formulaSettings.CalculationMode = CalcModeType.Manual;
                // Optional: prevent calculation on save while in manual mode
                formulaSettings.CalculateOnSave = false;

                // Reference to the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // -------------------------
                // Bulk data import starts
                // -------------------------

                // Example: import a large 2‑dimensional array of numeric values
                int rows = 10000;
                int cols = 10;
                object[,] data = new object[rows, cols];
                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        data[r, c] = r * cols + c + 1; // sample data
                    }
                }

                // Import the array starting at cell A1 using manual cell assignment
                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        cells[r, c].Value = data[r, c];
                    }
                }

                // Example: add some formulas that depend on the imported data
                // Sum of each column placed in the row after the data
                for (int c = 0; c < cols; c++)
                {
                    // Get column letters (e.g., "A", "AA")
                    string colLetter = CellIndexToName(0, c);
                    colLetter = System.Text.RegularExpressions.Regex.Replace(colLetter, @"\d", string.Empty);

                    string startAddr = $"{colLetter}1";
                    string endAddr = $"{colLetter}{rows}";
                    cells[rows, c].Formula = $"=SUM({startAddr}:{endAddr})";
                }

                // -------------------------
                // Bulk data import ends
                // -------------------------

                // Re‑enable automatic calculation (or set to desired mode)
                formulaSettings.CalculationMode = CalcModeType.Automatic;
                // Enable calculation on save if you want the file to be saved with calculated values
                formulaSettings.CalculateOnSave = true;

                // Recalculate all formulas now that data import is finished
                workbook.CalculateFormula();

                // Save the workbook (lifecycle save)
                workbook.Save("BulkImportResult.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Helper method to convert zero‑based row/column indexes to Excel cell name (e.g., 0,0 -> "A1")
        private static string CellIndexToName(int row, int column)
        {
            // Convert column index to letters
            string colName = "";
            int dividend = column + 1;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                colName = Convert.ToChar('A' + modulo) + colName;
                dividend = (dividend - modulo) / 26;
            }
            // Row index is zero‑based, Excel rows start at 1
            return $"{colName}{row + 1}";
        }
    }
}
