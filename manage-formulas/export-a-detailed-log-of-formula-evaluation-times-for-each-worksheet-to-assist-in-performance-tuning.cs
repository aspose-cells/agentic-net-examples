using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using Aspose.Cells;

// Custom monitor that records evaluation time for each cell and aggregates per worksheet
class FormulaEvaluationLogger : AbstractCalculationMonitor
{
    // Simple struct to identify a cell uniquely
    private struct CellKey : IEquatable<CellKey>
    {
        public int Sheet;
        public int Row;
        public int Column;
        public CellKey(int sheet, int row, int column)
        {
            Sheet = sheet;
            Row = row;
            Column = column;
        }
        public bool Equals(CellKey other) => Sheet == other.Sheet && Row == other.Row && Column == other.Column;
        public override int GetHashCode() => (Sheet, Row, Column).GetHashCode();
    }

    // Stopwatch per cell while it is being calculated
    private readonly Dictionary<CellKey, Stopwatch> _activeStops = new Dictionary<CellKey, Stopwatch>();

    // Log lines per worksheet
    private readonly Dictionary<int, List<string>> _cellLogs = new Dictionary<int, List<string>>();

    // Total time per worksheet (in milliseconds)
    private readonly Dictionary<int, double> _sheetTotals = new Dictionary<int, double>();

    // Called before a cell is evaluated
    public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
    {
        var key = new CellKey(sheetIndex, rowIndex, columnIndex);
        var sw = Stopwatch.StartNew();
        _activeStops[key] = sw;
    }

    // Called after a cell is evaluated
    public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
    {
        var key = new CellKey(sheetIndex, rowIndex, columnIndex);
        if (_activeStops.TryGetValue(key, out var sw))
        {
            sw.Stop();
            double ms = sw.Elapsed.TotalMilliseconds;

            // Record per‑cell detail
            if (!_cellLogs.ContainsKey(sheetIndex))
                _cellLogs[sheetIndex] = new List<string>();
            _cellLogs[sheetIndex].Add($"Cell ({rowIndex},{columnIndex}) evaluated in {ms:F3} ms");

            // Accumulate worksheet total
            if (!_sheetTotals.ContainsKey(sheetIndex))
                _sheetTotals[sheetIndex] = 0;
            _sheetTotals[sheetIndex] += ms;

            _activeStops.Remove(key);
        }
    }

    // Optional handling of circular references (use default behavior)
    public override bool OnCircular(IEnumerator circularCellsData)
    {
        return base.OnCircular(circularCellsData);
    }

    // Print a readable report to the console
    public void PrintReport()
    {
        foreach (var kvp in _sheetTotals)
        {
            int sheetIdx = kvp.Key;
            double totalMs = kvp.Value;
            Console.WriteLine($"Worksheet {sheetIdx} total formula evaluation time: {totalMs:F3} ms");
            if (_cellLogs.TryGetValue(sheetIdx, out var lines))
            {
                foreach (var line in lines)
                {
                    Console.WriteLine("  " + line);
                }
            }
        }
    }
}

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Set up calculation options with the custom monitor
        CalculationOptions calcOptions = new CalculationOptions();
        var monitor = new FormulaEvaluationLogger();
        calcOptions.CalculationMonitor = monitor;

        // Calculate all formulas in the workbook (recursive across worksheets)
        workbook.CalculateFormula(calcOptions);

        // Output the detailed evaluation log
        monitor.PrintReport();

        // Save the workbook after calculation if needed
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}