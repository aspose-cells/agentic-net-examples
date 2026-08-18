// Title: Measure per‑cell formula execution time with a custom AbstractCalculationMonitor in Aspose.Cells for .NET
// Description: This C# example shows how to profile each cell's formula evaluation by extending Aspose.Cells' AbstractCalculationMonitor. A Stopwatch is started in BeforeCalculate and stopped in AfterCalculate, logging elapsed milliseconds, original and calculated values, and change status. The monitor is attached via CalculationOptions and used with Cell.Calculate inside a processing loop before saving the workbook.
// Keywords: Aspose.Cells | C# | .NET | AbstractCalculationMonitor | Cell.Calculate | formula timing | performance profiling | Excel calculation benchmark | stopwatch | custom calculation monitor | GitHub example
// Common Searches: how to time individual cell calculations in Aspose.Cells | measure formula execution duration with AbstractCalculationMonitor C# | log per‑cell calculation latency using Aspose.Cells .NET | benchmark Excel formula performance with Aspose.Cells | profile cell.Calculate runtime in C#
// Developer Intent: Capture and display the execution time of each cell's formula evaluation during workbook processing.
// Use Cases: Identify and optimize slow‑running formulas in large spreadsheets. | Compare performance of volatile functions (e.g., NOW) against simple arithmetic. | Generate a detailed timing report for auditing calculation efficiency before saving the file.
// AI Prompts: Create a C# version of TimingCalculationMonitor that writes timing data to a CSV file instead of the console. | Show how to aggregate the logged timings into a summary table after all cells have been calculated. | Explain how to extend TimingCalculationMonitor to record memory usage alongside execution time for each cell.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Custom monitor that measures time taken for each cell calculation
    // This C# example shows how to profile each cell's formula evaluation by extending Aspose.Cells' AbstractCalculationMonitor. A Stopwatch is started in BeforeCalculate and stopped in AfterCalculate, logging elapsed milliseconds, original and calculated values, and change status. The monitor is attached via CalculationOptions and used with Cell.Calculate inside a processing loop before saving the workbook.
    public class TimingCalculationMonitor : AbstractCalculationMonitor
    {
        // Store start time for the current cell being calculated
        private readonly Stopwatch _stopwatch = new Stopwatch();

        public override void BeforeCalculate(int sheetIndex, int rowIndex, int colIndex)
        {
            // Restart stopwatch before each cell calculation
            _stopwatch.Restart();
        }

        public override void AfterCalculate(int sheetIndex, int rowIndex, int colIndex)
        {
            // Stop stopwatch and log elapsed time
            _stopwatch.Stop();
            TimeSpan elapsed = _stopwatch.Elapsed;
            Console.WriteLine($"Calculated cell (Sheet {sheetIndex}) Row {rowIndex}, Column {colIndex} in {elapsed.TotalMilliseconds} ms");
            Console.WriteLine($"  Original Value: {OriginalValue}, Calculated Value: {CalculatedValue}, Value Changed: {ValueChanged}");
            Console.WriteLine(new string('-', 50));
        }
    }

    public class CellCalculationTimingDemo
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data and formulas
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].Formula = "=A1+A2";               // Simple addition
            cells["B1"].Formula = "=SUM(A1:A3)";          // Sum range
            cells["B2"].Formula = "=NOW()";              // Volatile function
            cells["C1"].Formula = "=IF(A1>5,\"High\",\"Low\")";

            // Prepare calculation options with the custom timing monitor
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = new TimingCalculationMonitor(),
                Recursive = true,
                IgnoreError = false
            };

            // List of cells to calculate individually
            List<Cell> cellsToCalculate = new List<Cell>
            {
                cells["A3"],
                cells["B1"],
                cells["B2"],
                cells["C1"]
            };

            // Calculate each cell separately, timing will be logged by the monitor
            foreach (Cell cell in cellsToCalculate)
            {
                cell.Calculate(options);
            }

            // Optionally, calculate the whole workbook at once (timings will also be logged)
            // workbook.CalculateFormula(options);

            // Save the workbook
            workbook.Save("CellCalculationTimingDemo.xlsx");
        }
    }

    // Entry point
    class Program
    {
        static void Main()
        {
            CellCalculationTimingDemo.Run();
        }
    }
}
