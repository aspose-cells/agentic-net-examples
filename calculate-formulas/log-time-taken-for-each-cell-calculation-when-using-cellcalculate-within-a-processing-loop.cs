using System;
using System.Collections.Generic;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsTimingDemo
{
    // Custom monitor that measures the time taken to calculate each cell
    public class TimingCalculationMonitor : AbstractCalculationMonitor
    {
        // Stopwatch is reused for each cell calculation
        private readonly Stopwatch _stopwatch = new Stopwatch();

        // Called before a cell is calculated
        public override void BeforeCalculate(int sheetIndex, int rowIndex, int colIndex)
        {
            _stopwatch.Restart(); // start timing
        }

        // Called after a cell is calculated
        public override void AfterCalculate(int sheetIndex, int rowIndex, int colIndex)
        {
            _stopwatch.Stop(); // stop timing
            TimeSpan elapsed = _stopwatch.Elapsed;

            // Log the timing information together with cell address
            Console.WriteLine(
                $"Calculated Sheet{sheetIndex} Cell[{rowIndex}, {colIndex}] " +
                $"in {elapsed.TotalMilliseconds:F3} ms. " +
                $"ValueChanged: {ValueChanged}, " +
                $"Original: {OriginalValue}, New: {CalculatedValue}");
        }
    }

    public class Program
    {
        public static void Main()
        {
            // -------------------------------------------------
            // 1. Create a new workbook and add sample formulas
            // -------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].Formula = "=A1+A2";          // simple addition
            cells["B1"].Formula = "=A3*2";           // dependent on A3
            cells["C1"].Formula = "=NOW()";          // volatile function
            cells["D1"].Formula = "=SUM(A1:A2)";     // aggregate function

            // -------------------------------------------------
            // 2. Prepare calculation options with the timing monitor
            // -------------------------------------------------
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = new TimingCalculationMonitor(),
                // Ensure recursive calculation so dependent cells are processed
                Recursive = true
            };

            // -------------------------------------------------
            // 3. Loop through the cells that contain formulas and calculate each one individually
            // -------------------------------------------------
            // Collect cells that have formulas
            List<Cell> formulaCells = new List<Cell>();
            foreach (Cell cell in cells)
            {
                if (cell.IsFormula)
                {
                    formulaCells.Add(cell);
                }
            }

            // Calculate each formula cell using Cell.Calculate with the same options
            foreach (Cell cell in formulaCells)
            {
                cell.Calculate(options);
            }

            // -------------------------------------------------
            // 4. Save the workbook (lifecycle rule: use provided save logic)
            // -------------------------------------------------
            workbook.Save("TimingCalculationDemo.xlsx");
        }
    }
}