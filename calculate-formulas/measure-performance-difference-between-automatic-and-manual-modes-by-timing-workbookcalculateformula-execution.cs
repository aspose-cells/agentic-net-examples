using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsPerformanceDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate the worksheet with sample data and formulas
            // Column A will contain base values, columns B‑J will contain formulas that reference the previous column
            const int totalRows = 500;
            const int totalCols = 10; // A‑J

            for (int row = 0; row < totalRows; row++)
            {
                // Base value in column A
                cells[row, 0].PutValue(row + 1);

                // Formulas in columns B‑J
                for (int col = 1; col < totalCols; col++)
                {
                    // Example formula: =B1*2 (each cell doubles the value of the left neighbour)
                    string leftCellName = CellsHelper.CellIndexToName(row, col - 1);
                    cells[row, col].Formula = $"={leftCellName}*2";
                }
            }

            // -----------------------------------------------------------------
            // Measure calculation time in Automatic mode
            // -----------------------------------------------------------------
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

            Stopwatch swAuto = Stopwatch.StartNew();
            workbook.CalculateFormula(); // calculate all formulas
            swAuto.Stop();

            Console.WriteLine($"Automatic mode calculation time: {swAuto.ElapsedMilliseconds} ms");

            // -----------------------------------------------------------------
            // Measure calculation time in Manual mode
            // -----------------------------------------------------------------
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

            // Optionally change a cell to simulate a typical manual recalculation scenario
            cells[0, 0].PutValue(999);

            Stopwatch swManual = Stopwatch.StartNew();
            workbook.CalculateFormula(); // explicit calculation required in Manual mode
            swManual.Stop();

            Console.WriteLine($"Manual mode calculation time: {swManual.ElapsedMilliseconds} ms");

            // Save the workbook (optional, demonstrates usage of the save API)
            workbook.Save("PerformanceDemo.xlsx");
        }
    }
}