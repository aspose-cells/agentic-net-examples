using System;
using System.Collections.Generic;
using System.Diagnostics;
using Aspose.Cells;

class CellCalculationTimer : AbstractCalculationMonitor
{
    // Store a stopwatch for each cell being calculated
    private readonly Dictionary<string, Stopwatch> _timers = new Dictionary<string, Stopwatch>();

    private string GetKey(int sheet, int row, int col) => $"{sheet}:{row}:{col}";

    public override void BeforeCalculate(int sheetIndex, int rowIndex, int colIndex)
    {
        // Start timing before the cell is calculated
        var key = GetKey(sheetIndex, rowIndex, colIndex);
        var sw = Stopwatch.StartNew();
        _timers[key] = sw;
    }

    public override void AfterCalculate(int sheetIndex, int rowIndex, int colIndex)
    {
        // Stop timing after calculation and log the elapsed time
        var key = GetKey(sheetIndex, rowIndex, colIndex);
        if (_timers.TryGetValue(key, out var sw))
        {
            sw.Stop();
            Console.WriteLine($"Cell (Sheet {sheetIndex}, Row {rowIndex}, Column {colIndex}) calculated in {sw.ElapsedMilliseconds} ms. New Value: {CalculatedValue}");
            _timers.Remove(key);
        }
    }
}

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data and formulas
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);
        sheet.Cells["A3"].Formula = "=A1+A2";
        sheet.Cells["B1"].Formula = "=SUM(A1:A3)";
        sheet.Cells["C1"].Formula = "=NOW()"; // volatile function

        // Create a calculation monitor that logs timing
        var monitor = new CellCalculationTimer();

        // Set calculation options with the custom monitor
        var options = new CalculationOptions { CalculationMonitor = monitor };

        // Loop through all cells, calculate only those that contain formulas
        foreach (Cell cell in sheet.Cells)
        {
            if (cell.IsFormula)
            {
                cell.Calculate(options);
            }
        }

        // Save the workbook
        workbook.Save("TimedCalculations.xlsx");
    }
}