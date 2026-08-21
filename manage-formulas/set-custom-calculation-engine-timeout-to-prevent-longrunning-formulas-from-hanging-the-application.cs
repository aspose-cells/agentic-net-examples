// Title: Configure a Calculation Timeout in Aspose.Cells for .NET to Abort Long‑Running Formulas
// Description: Learn how to create a custom TimeoutCalculationMonitor by extending AbstractCalculationMonitor, attach it to CalculationOptions, and use wb.CalculateFormula with a time limit. The monitor throws a TimeoutException when the elapsed time exceeds the defined threshold, preventing the application from hanging during heavy spreadsheet calculations.
// Keywords: Aspose.Cells calculation timeout | C# custom calculation monitor | AbstractCalculationMonitor example | prevent long‑running formulas | CalculationOptions timeout | Workbook.CalculateFormula limit | TimeoutException in Aspose.Cells | Excel engine performance guard
// Common Searches: Aspose.Cells set formula calculation timeout C# | How to stop long calculations in Aspose.Cells | Custom AbstractCalculationMonitor for timeout | Limit workbook.CalculateFormula execution time | Abort hanging formulas Aspose.Cells .NET
// Developer Intent: Implement a time‑bound calculation engine so that any formula exceeding the specified duration is automatically cancelled.
// Use Cases: Avoid UI freezes in desktop apps by aborting formulas that run longer than a few seconds. | Protect server resources when processing user‑uploaded spreadsheets by enforcing a maximum calculation time. | Detect and log excessively complex or recursive formulas by catching the timeout exception. | Provide a graceful fallback (e.g., display a warning) when a workbook cannot be fully calculated within the allowed time.
// AI Prompts: Generate a TimeoutCalculationMonitor that records the cell address before throwing a TimeoutException. | Show how to retry workbook.CalculateFormula with an increased timeout after the first attempt fails. | Create code that logs the sheet name, row, and column of the cell that triggered the timeout.

using System;
using System.Diagnostics;
using Aspose.Cells;

// Custom calculation monitor that aborts calculation after a specified time limit
// Learn how to create a custom TimeoutCalculationMonitor by extending AbstractCalculationMonitor, attach it to CalculationOptions, and use wb.CalculateFormula with a time limit. The monitor throws a TimeoutException when the elapsed time exceeds the defined threshold, preventing the application from hanging during heavy spreadsheet calculations.
class TimeoutCalculationMonitor : AbstractCalculationMonitor
{
    private readonly Stopwatch _stopwatch;
    private readonly long _maxMilliseconds;

    public TimeoutCalculationMonitor(int maxMilliseconds)
    {
        _maxMilliseconds = maxMilliseconds;
        _stopwatch = new Stopwatch();
        _stopwatch.Start();
    }

    // Called before each cell calculation
    public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
    {
        if (_stopwatch.ElapsedMilliseconds > _maxMilliseconds)
        {
            // Stop the calculation by throwing an exception
            throw new TimeoutException($"Formula calculation exceeded the time limit of {_maxMilliseconds} ms.");
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
        // Create a new workbook
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate cells with data that could cause long‑running calculations
        cells["A1"].PutValue(1);
        cells["A2"].Formula = "=A1+1";
        // Example of a potentially heavy formula (recursive sum)
        cells["B1"].Formula = "=SUM(A1:A1000)";

        // Configure calculation options with the timeout monitor (e.g., 2 seconds)
        CalculationOptions calcOptions = new CalculationOptions
        {
            CalculationMonitor = new TimeoutCalculationMonitor(2000) // 2000 ms = 2 seconds
        };

        try
        {
            // Perform calculation using the custom options
            wb.CalculateFormula(calcOptions);
            Console.WriteLine("Calculation completed successfully.");
        }
        catch (TimeoutException ex)
        {
            Console.WriteLine("Calculation aborted: " + ex.Message);
        }

        // Save the workbook (standard lifecycle)
        wb.Save("TimeoutDemo.xlsx");
    }
}
