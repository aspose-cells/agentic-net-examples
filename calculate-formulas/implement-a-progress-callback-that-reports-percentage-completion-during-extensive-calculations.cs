// Title: Aspose.Cells C# – Custom Calculation Monitor for Percentage Progress
// Description: Demonstrates how to extend AbstractCalculationMonitor to create a ProgressCalculationMonitor that counts processed cells, calculates completion percentage, and writes it to the console during workbook.CalculateFormula. The example fills a worksheet with many formulas, attaches the monitor via CalculationOptions, runs the calculation, and saves the result.
// Keywords: Aspose.Cells | C# | AbstractCalculationMonitor | CalculationOptions | formula calculation progress | percentage completion callback | custom calculation monitor | large worksheet performance
// Common Searches: Aspose.Cells progress callback C# | how to monitor formula calculation percentage | custom AbstractCalculationMonitor example | track workbook calculation progress Aspose.Cells | display calculation status in .NET
// Developer Intent: Add a callback that reports the percentage of formulas processed during a workbook calculation.
// Use Cases: Log real‑time progress of formula evaluation in a massive sheet. | Update a UI progress bar while CalculateFormula runs. | Provide feedback in console applications for long‑running calculations. | Integrate custom monitoring with parallel calculation settings.
// AI Prompts: Generate a C# ProgressCalculationMonitor that updates a WinForms ProgressBar instead of writing to the console. | Show how to throttle console output in AfterCalculate to improve performance for thousands of cells. | Create an example that combines a custom calculation monitor with Aspose.Cells parallel calculation mode while preserving overall progress reporting.

using System;
using System.Collections;
using Aspose.Cells;

// Demonstrates how to extend AbstractCalculationMonitor to create a ProgressCalculationMonitor that counts processed cells, calculates completion percentage, and writes it to the console during workbook.CalculateFormula. The example fills a worksheet with many formulas, attaches the monitor via CalculationOptions, runs the calculation, and saves the result.
class ProgressCalculationMonitor : AbstractCalculationMonitor
{
    private readonly int _totalCells;
    private int _processedCells;

    public ProgressCalculationMonitor(int totalCells)
    {
        _totalCells = totalCells;
        _processedCells = 0;
    }

    // Called after each cell is calculated
    public override void AfterCalculate(int sheetIndex, int rowIndex, int colIndex)
    {
        _processedCells++;
        double percent = (double)_processedCells / _totalCells * 100;
        Console.WriteLine($"Calculated cell [Sheet {sheetIndex}, Row {rowIndex}, Column {colIndex}] - {percent:F2}% completed");
    }

    // Optional: before calculation (not used here)
    public override void BeforeCalculate(int sheetIndex, int rowIndex, int colIndex) { }

    // Optional: handle circular references (continue calculation)
    public override bool OnCircular(IEnumerator circularCellsData) => true;
}

class CalculationProgressDemo
{
    public static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate the sheet with many formulas to simulate an extensive calculation
        int rowCount = 500; // adjust as needed for testing
        for (int i = 0; i < rowCount; i++)
        {
            // Simple numeric value
            worksheet.Cells[i, 0].PutValue(i + 1);
            // Formula that depends on the numeric value
            worksheet.Cells[i, 1].Formula = $"=A{i + 1}*2";
        }

        // Total number of formula cells that will be processed
        int totalFormulaCells = rowCount;

        // Set up calculation options with the custom progress monitor
        CalculationOptions calcOptions = new CalculationOptions
        {
            CalculationMonitor = new ProgressCalculationMonitor(totalFormulaCells)
        };

        // Perform the calculation with monitoring
        workbook.CalculateFormula(calcOptions);

        // Save the workbook to verify results
        workbook.Save("CalculationProgressDemo.xlsx");
    }
}
