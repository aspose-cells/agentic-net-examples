using System;
using Aspose.Cells;

// Custom calculation monitor that aborts calculation after a time limit (in milliseconds)
class TimeLimitMonitor : AbstractCalculationMonitor
{
    private readonly long _timeLimitMs;
    private readonly DateTime _startTime;

    public TimeLimitMonitor(long timeLimitMs)
    {
        _timeLimitMs = timeLimitMs;
        _startTime = DateTime.UtcNow;
    }

    // Called before each cell calculation
    public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
    {
        if ((DateTime.UtcNow - _startTime).TotalMilliseconds > _timeLimitMs)
        {
            // Abort the whole calculation by throwing an exception
            throw new TimeoutException($"Formula calculation exceeded the time limit of {_timeLimitMs} ms.");
        }
    }

    // Called after each cell calculation (no action needed)
    public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex) { }

    // Called when a circular reference is detected (default handling)
    public override bool OnCircular(System.Collections.IEnumerator circularCellsData)
    {
        return base.OnCircular(circularCellsData);
    }
}

class SetCalculationTimeoutDemo
{
    static void Main()
    {
        // Create a new workbook and add some formulas that could take time
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        Cells cells = ws.Cells;

        // Example of a potentially long‑running formula (e.g., iterative calculation)
        cells["A1"].Formula = "=SUMPRODUCT(ROW(INDIRECT(\"1:1000000\")),COLUMN(INDIRECT(\"A:Z\")))";

        // Set up calculation options with the custom timeout monitor (e.g., 2 seconds)
        CalculationOptions calcOptions = new CalculationOptions
        {
            CalculationMonitor = new TimeLimitMonitor(2000) // 2000 ms = 2 seconds
        };

        try
        {
            // Perform calculation with the timeout monitor
            wb.CalculateFormula(calcOptions);
            Console.WriteLine("Calculation completed successfully.");
            Console.WriteLine("Result in A1: " + cells["A1"].Value);
        }
        catch (TimeoutException ex)
        {
            Console.WriteLine("Calculation aborted: " + ex.Message);
        }

        // Save the workbook (optional)
        wb.Save("TimeoutDemo.xlsx");
    }
}