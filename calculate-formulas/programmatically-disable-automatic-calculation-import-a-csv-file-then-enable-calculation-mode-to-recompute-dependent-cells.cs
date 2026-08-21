// Title: Aspose.Cells .NET: Disable Auto‑Calc, Import CSV, Re‑enable Calculation and Recalculate Formulas (C#)
// Description: Demonstrates how to set a workbook's CalculationMode to Manual before importing CSV data with ImportCSV, then switch back to Automatic (or AutomaticExceptTable) and force a full formula recalculation using CalculateFormula, finally saving the result as XLSX. This approach prevents unnecessary formula evaluation during bulk data loads and improves performance.
// Keywords: Aspose.Cells C# | disable automatic calculation | ImportCSV Aspose.Cells | re‑enable calculation mode | CalculateFormula | CSV to Excel conversion | performance optimization Aspose.Cells | CalcModeType Manual | CalcModeType Automatic
// Common Searches: how to turn off formula calculation in Aspose.Cells before importing CSV | Aspose.Cells import CSV without triggering formulas | recalculate all formulas after CSV import Aspose.Cells .NET | set calculation mode manual then automatic Aspose.Cells | force formula evaluation after data load Aspose.Cells
// Developer Intent: Load CSV data without triggering formulas, then enable calculation and recompute all dependent cells.
// Use Cases: Bulk import of financial data from CSV into a model while avoiding intermediate formula runs. | Generating large reports where raw data is loaded first and formulas are evaluated once at the end. | Improving performance of data‑intensive workbooks by toggling calculation mode around ImportCSV.
// AI Prompts: Write C# code with Aspose.Cells that disables automatic calculation, imports a CSV file, re‑enables calculation, forces formula recalculation, and saves the workbook. | Explain the performance benefits of setting CalcModeType to Manual before ImportCSV and how to restore automatic calculation afterward. | Provide a concise Aspose.Cells example that toggles CalcModeType, uses ImportCSV, calls CalculateFormula, and outputs an XLSX file.

using System;
using Aspose.Cells;

namespace AsposeCellsCalcModeExample
{
    // Demonstrates how to set a workbook's CalculationMode to Manual before importing CSV data with ImportCSV, then switch back to Automatic (or AutomaticExceptTable) and force a full formula recalculation using CalculateFormula, finally saving the result as XLSX. This approach prevents unnecessary formula evaluation during bulk data loads and improves performance.
    class Program
    {
        static void Main(string[] args)
        {
            // 1. Create a new workbook (empty)
            Workbook workbook = new Workbook();

            // 2. Access the first worksheet and its cells collection
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // 3. Disable automatic calculation before importing data
            //    This prevents any formulas from being evaluated while the CSV is being loaded.
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

            // 4. Import CSV data
            //    Adjust the file path, delimiter and other parameters as needed.
            string csvPath = "input.csv";          // Path to your CSV file
            string delimiter = ",";                // CSV delimiter
            bool convertNumeric = true;            // Convert numeric strings to numbers
            int startRow = 0;                      // Zero‑based index for the first row (A1)
            int startColumn = 0;                   // Zero‑based index for the first column (A1)

            cells.ImportCSV(csvPath, delimiter, convertNumeric, startRow, startColumn);

            // 5. Re‑enable automatic calculation (or set to AutomaticExceptTable if preferred)
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

            // 6. Force calculation of all formulas now that the data is loaded
            workbook.CalculateFormula();

            // 7. Save the workbook to verify results
            workbook.Save("output.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("CSV imported and formulas recalculated successfully.");
        }
    }
}
