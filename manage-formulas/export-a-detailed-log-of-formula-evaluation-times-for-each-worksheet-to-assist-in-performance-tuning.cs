// Title: Log per‑worksheet formula calculation times with Aspose.Cells for .NET
// Description: Demonstrates how to create a custom AbstractCalculationMonitor that uses a Stopwatch to record the duration of each cell's formula evaluation, aggregates the times per worksheet, and prints the results for performance tuning.
// Keywords: Aspose.Cells | C# | .NET | formula calculation timing | custom calculation monitor | AbstractCalculationMonitor | performance profiling | Excel formula benchmark | worksheet calculation time
// Common Searches: Aspose.Cells log formula calculation time per sheet | measure Excel formula performance with C# | custom calculation monitor example Aspose.Cells | export worksheet formula evaluation timings | profile Excel workbook calculation speed .NET
// Developer Intent: Capture and export the execution time of formula calculations for each worksheet in an Excel workbook.
// Use Cases: Identify worksheets that cause slow recalculation and target them for optimization. | Generate a performance report of formula evaluation times before releasing a workbook. | Compare calculation speed after applying formula or setting changes across multiple workbooks.
// AI Prompts: Write a C# method that saves the sheet timing dictionary to a CSV file for further analysis. | Extend FormulaTimingMonitor to also count the number of formulas evaluated per worksheet. | Create a script that runs the timing monitor on a batch of workbooks and aggregates overall statistics.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using Aspose.Cells;

// Custom monitor to capture calculation time per cell and aggregate per worksheet
// Demonstrates how to create a custom AbstractCalculationMonitor that uses a Stopwatch to record the duration of each cell's formula evaluation, aggregates the times per worksheet, and prints the results for performance tuning.
class FormulaTimingMonitor : AbstractCalculationMonitor
{
    // Stores elapsed time for each worksheet (key = sheet index)
    private readonly Dictionary<int, TimeSpan> _sheetTimes = new Dictionary<int, TimeSpan>();

    // Temporary stopwatch for the cell currently being calculated
    private readonly Dictionary<string, Stopwatch> _activeStops = new Dictionary<string, Stopwatch>();

    // Called before a cell is calculated
    public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
    {
        string key = $"{sheetIndex}_{rowIndex}_{columnIndex}";
        var sw = new Stopwatch();
        sw.Start();
        _activeStops[key] = sw;
    }

    // Called after a cell is calculated
    public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
    {
        string key = $"{sheetIndex}_{rowIndex}_{columnIndex}";
        if (_activeStops.TryGetValue(key, out var sw))
        {
            sw.Stop();
            if (!_sheetTimes.ContainsKey(sheetIndex))
                _sheetTimes[sheetIndex] = TimeSpan.Zero;
            _sheetTimes[sheetIndex] += sw.Elapsed;
            _activeStops.Remove(key);
        }
    }

    // Expose the collected timings
    public IReadOnlyDictionary<int, TimeSpan> SheetTimes => _sheetTimes;
}

class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Set up calculation options with the custom monitor
        CalculationOptions calcOptions = new CalculationOptions();
        var monitor = new FormulaTimingMonitor();
        calcOptions.CalculationMonitor = monitor;

        // Perform full formula calculation using the monitor
        workbook.CalculateFormula(calcOptions);

        // Output detailed timing information per worksheet
        Console.WriteLine("Formula evaluation times per worksheet:");
        foreach (var kvp in monitor.SheetTimes)
        {
            string sheetName = workbook.Worksheets[kvp.Key].Name;
            Console.WriteLine($"- Sheet \"{sheetName}\" (Index {kvp.Key}): {kvp.Value.TotalMilliseconds} ms");
        }

        // Save the workbook if further processing is required
        workbook.Save("output.xlsx");
    }
}
