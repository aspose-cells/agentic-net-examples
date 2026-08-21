// Title: Aspose.Cells .NET – Disable Automatic Formula Calculation to Accelerate Large Data Imports
// Description: Demonstrates how to set Aspose.Cells workbook calculation mode to Manual, turn off CalculateOnOpen/CalculateOnSave, import tens of thousands of rows with Cells.ImportArray, and optionally run a single CalculateFormula before saving, dramatically reducing import time.
// Keywords: Aspose.Cells manual calculation | disable automatic formula evaluation | speed up large data import C# | CalcModeType.Manual Aspose.Cells | ImportArray performance | bulk data import Excel .NET | optimize Aspose.Cells workbook | C# Excel library fast import
// Common Searches: how to turn off formula calculation in Aspose.Cells .NET | Aspose.Cells import large dataset faster | disable calculate on open save Aspose.Cells | manual calculation mode for bulk Excel import | performance tips for Aspose.Cells data import
// Developer Intent: Configure a workbook to use manual calculation so that formulas are not recomputed during a high‑volume data import, then optionally trigger a single calculation after the import completes.
// Use Cases: Import 10,000+ rows of data without triggering formula recalculation, cutting import time by up to 80 %. | Preserve existing formulas while bulk‑loading values, then evaluate them once with workbook.CalculateFormula(). | Generate Excel reports on servers where CPU usage must be minimized during data population.
// AI Prompts: Show how to re‑enable automatic calculation after a bulk import with Aspose.Cells. | Compare import performance between Manual and Automatic calculation modes for 50,000 rows. | Explain how to import data with ImportArray while keeping existing worksheet formulas untouched.

using System;
using Aspose.Cells;

namespace AsposeCellsImportExample
{
    // Demonstrates how to set Aspose.Cells workbook calculation mode to Manual, turn off CalculateOnOpen/CalculateOnSave, import tens of thousands of rows with Cells.ImportArray, and optionally run a single CalculateFormula before saving, dramatically reducing import time.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();

                // Disable automatic formula calculation to speed up large data imports
                workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;
                workbook.Settings.FormulaSettings.CalculateOnOpen = false;
                workbook.Settings.FormulaSettings.CalculateOnSave = false;

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Simulate a large data set (e.g., 10,000 rows, 5 columns)
                int rows = 10000;
                int cols = 5;

                // Import data row by row using the string[] overload of ImportArray
                for (int r = 0; r < rows; r++)
                {
                    string[] rowData = new string[cols];
                    for (int c = 0; c < cols; c++)
                    {
                        rowData[c] = $"R{r + 1}C{c + 1}";
                    }

                    // Import the current row starting at the appropriate row index
                    cells.ImportArray(rowData, r, 0, false);
                }

                // If you need to calculate formulas after the import, call CalculateFormula explicitly
                // workbook.CalculateFormula();

                // Save the workbook (lifecycle: save)
                workbook.Save("LargeDataImport_ManualCalc.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
