// Title: Log per‑cell formula calculation time with a custom CalculationMonitor using Cell.Calculate in Aspose.Cells for .NET
// AI Prompts: Create a subclass of AbstractCalculationMonitor that starts a Stopwatch in BeforeCalculate and writes the elapsed milliseconds together with the cell address in AfterCalculate. | Iterate over the worksheet’s used range, invoke Cell.Calculate on each formula cell with CalculationOptions that reference the custom monitor, and output the timing information. | Store each cell’s duration in a collection for later analysis or reporting, then save the workbook.
// Common Searches: how to profile individual cell formula evaluation in Aspose.Cells C# | Aspose.Cells custom calculation monitor example for measuring performance | log execution time of each formula cell using Cell.Calculate .NET | track per‑cell calculation duration in an Excel workbook with Aspose.Cells | measure formula recalculation latency in Aspose.Cells for .NET
// Tags: Cell.Calculate performance monitoring | Aspose.Cells custom AbstractCalculationMonitor | per‑cell formula timing .NET | measure Excel calculation duration Aspose | stopwatch timing for cell evaluation

using System;
using System.Diagnostics;
using Aspose.Cells;

// The example creates a workbook, defines a TimingCalculationMonitor that uses a Stopwatch to record and print the elapsed milliseconds for each formula cell during calculation, configures CalculationOptions to use this monitor, loops through all used cells calling Cell.Calculate with those options, and finally saves the workbook.
class CellCalculationTimer
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data and formulas
        worksheet.Cells["A1"].PutValue(10);
        worksheet.Cells["A2"].PutValue(20);
        worksheet.Cells["A3"].Formula = "=A1+A2";
        worksheet.Cells["B1"].Formula = "=A3*2";
        worksheet.Cells["C1"].Formula = "=NOW()"; // volatile function

        // Create a custom calculation monitor that measures time per cell
        var monitor = new TimingCalculationMonitor();

        // Set calculation options with the monitor
        CalculationOptions options = new CalculationOptions
        {
            CalculationMonitor = monitor,
            Recursive = true
        };

        // Iterate over all used cells and calculate each formula individually
        var usedRange = worksheet.Cells.MaxDisplayRange;
        for (int row = usedRange.FirstRow; row <= usedRange.FirstRow + usedRange.RowCount - 1; row++)
        {
            for (int col = usedRange.FirstColumn; col <= usedRange.FirstColumn + usedRange.ColumnCount - 1; col++)
            {
                Cell cell = worksheet.Cells[row, col];
                if (cell.IsFormula)
                {
                    // Calculate the cell using the options that contain our monitor
                    cell.Calculate(options);
                }
            }
        }

        // Save the workbook
        workbook.Save("TimedCalculations.xlsx");
    }

    // Custom monitor that logs the elapsed time for each cell calculation
    private class TimingCalculationMonitor : AbstractCalculationMonitor
    {
        private Stopwatch _stopwatch = new Stopwatch();

        public override void BeforeCalculate(int sheetIndex, int rowIndex, int colIndex)
        {
            // Start timing before the cell is calculated
            _stopwatch.Restart();
        }

        public override void AfterCalculate(int sheetIndex, int rowIndex, int colIndex)
        {
            // Stop timing after calculation and output the result
            _stopwatch.Stop();
            Console.WriteLine($"Calculated cell (Sheet {sheetIndex}, Row {rowIndex}, Column {colIndex}) in {_stopwatch.Elapsed.TotalMilliseconds} ms. New Value: {CalculatedValue}");
        }
    }
}
