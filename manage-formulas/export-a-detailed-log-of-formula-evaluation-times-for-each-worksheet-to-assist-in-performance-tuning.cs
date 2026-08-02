// Title: Log formula evaluation times per worksheet with Aspose.Cells for .NET
// Description: Demonstrates how to create a custom TimingMonitor by extending AbstractCalculationMonitor to capture start and end timestamps for each cell, aggregate total duration and cell count per worksheet, and output a detailed performance log after Workbook.CalculateFormula. Ideal for tuning Excel formula calculation speed in .NET applications.
// Keywords: Aspose.Cells | C# formula performance | custom calculation monitor | worksheet calculation timing | Excel formula evaluation log | performance tuning Aspose.Cells | .NET workbook calculation | measure cell compute time
// Common Searches: Aspose.Cells track formula calculation time | How to log worksheet evaluation duration in C# | Custom calculation monitor example Aspose.Cells | Export formula performance data from Excel workbook | Measure average cell calculation time Aspose
// Developer Intent: Generate a per‑worksheet log that reports total, average, and count of formula evaluations to aid performance analysis.
// Use Cases: Identify sheets that dominate calculation time and target them for optimization. | Compare evaluation metrics before and after formula refactoring to verify speed gains. | Integrate timing logs into CI pipelines to detect regression in calculation performance.
// AI Prompts: Write code that saves the TimingMonitor sheet durations and cell counts to a CSV file. | Enhance TimingMonitor to record the maximum single‑cell calculation time per worksheet. | Create a unit test that verifies TimingMonitor aggregates total duration and cell count correctly for a sample workbook.

using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

namespace FormulaEvaluationLogDemo
{
    // Custom monitor to track calculation time per worksheet
    // Demonstrates how to create a custom TimingMonitor by extending AbstractCalculationMonitor to capture start and end timestamps for each cell, aggregate total duration and cell count per worksheet, and output a detailed performance log after Workbook.CalculateFormula. Ideal for tuning Excel formula calculation speed in .NET applications.
    public class TimingMonitor : AbstractCalculationMonitor
    {
        // Store start time for each cell (key: sheetIndex|row|col)
        private readonly Dictionary<string, DateTime> _startTimes = new Dictionary<string, DateTime>();

        // Aggregate timing data per worksheet
        private readonly Dictionary<int, TimeSpan> _sheetDurations = new Dictionary<int, TimeSpan>();
        private readonly Dictionary<int, int> _sheetCellCounts = new Dictionary<int, int>();

        // Called before a cell is calculated
        public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            string key = $"{sheetIndex}|{rowIndex}|{columnIndex}";
            _startTimes[key] = DateTime.UtcNow;
        }

        // Called after a cell is calculated
        public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            string key = $"{sheetIndex}|{rowIndex}|{columnIndex}";
            if (_startTimes.TryGetValue(key, out DateTime start))
            {
                TimeSpan elapsed = DateTime.UtcNow - start;

                // Accumulate elapsed time for the worksheet
                if (_sheetDurations.ContainsKey(sheetIndex))
                {
                    _sheetDurations[sheetIndex] += elapsed;
                    _sheetCellCounts[sheetIndex] += 1;
                }
                else
                {
                    _sheetDurations[sheetIndex] = elapsed;
                    _sheetCellCounts[sheetIndex] = 1;
                }

                _startTimes.Remove(key);
            }
        }

        // Called when a circular reference is detected (optional handling)
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            // Let the base implementation handle it
            return base.OnCircular(circularCellsData);
        }

        // Expose the collected timing data
        public IReadOnlyDictionary<int, TimeSpan> SheetDurations => _sheetDurations;
        public IReadOnlyDictionary<int, int> SheetCellCounts => _sheetCellCounts;
    }

    class Program
    {
        static void Main()
        {
            // Path to the workbook to be analyzed
            string inputPath = "input.xlsx";

            // Load the workbook (uses the standard load rule)
            Workbook workbook = new Workbook(inputPath);

            // Prepare calculation options with the custom monitor
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = new TimingMonitor()
            };

            // Perform full calculation (recursive = true for cross‑sheet dependencies)
            workbook.CalculateFormula(options);

            // Retrieve the monitor to extract timing information
            TimingMonitor monitor = (TimingMonitor)options.CalculationMonitor;

            // Output detailed log per worksheet
            Console.WriteLine("Formula Evaluation Time Log:");
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                int index = sheet.Index;
                TimeSpan totalTime = monitor.SheetDurations.ContainsKey(index) ? monitor.SheetDurations[index] : TimeSpan.Zero;
                int cellCount = monitor.SheetCellCounts.ContainsKey(index) ? monitor.SheetCellCounts[index] : 0;

                Console.WriteLine($"Worksheet: \"{sheet.Name}\" (Index {index})");
                Console.WriteLine($"  Cells evaluated : {cellCount}");
                Console.WriteLine($"  Total evaluation time : {totalTime.TotalMilliseconds} ms");
                if (cellCount > 0)
                {
                    Console.WriteLine($"  Average time per cell : {totalTime.TotalMilliseconds / cellCount:F4} ms");
                }
                Console.WriteLine();
            }

            // Optionally save the workbook after calculation (uses the standard save rule)
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}
