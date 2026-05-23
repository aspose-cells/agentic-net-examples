using System;
using Aspose.Cells;

namespace AsposeCellsCalcModeExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and its cells
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Disable automatic calculation by setting the mode to Manual
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

            // Import CSV data starting at cell A1 (row 0, column 0)
            // Adjust the file path, delimiter and conversion options as needed
            string csvPath = "data.csv";          // Path to your CSV file
            string delimiter = ",";               // CSV delimiter
            bool convertNumeric = true;           // Convert numeric strings to numbers
            int startRow = 0;                     // Zero‑based row index
            int startColumn = 0;                  // Zero‑based column index

            cells.ImportCSV(csvPath, delimiter, convertNumeric, startRow, startColumn);

            // Re‑enable automatic calculation (or choose another mode)
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the workbook to verify results
            workbook.Save("Result.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("CSV imported and formulas recalculated successfully.");
        }
    }
}