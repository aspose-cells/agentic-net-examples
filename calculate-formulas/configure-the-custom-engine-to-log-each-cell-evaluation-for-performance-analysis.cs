// Title: Log each cell evaluation with a custom CalculationMonitor in Aspose.Cells for .NET
// Description: Shows how to create a workbook, insert values and formulas, and attach a custom LoggingMonitor (derived from AbstractCalculationMonitor) via CalculationOptions. The monitor records before/after events and circular‑reference alerts during workbook.CalculateFormula, then saves the workbook.
// Keywords: Aspose.Cells | CalculationMonitor | CalculationOptions | log cell evaluation | formula performance profiling | circular reference handling | C# | .NET | custom monitor | performance analysis
// Common Searches: Aspose.Cells custom CalculationMonitor example | log each cell calculation Aspose.Cells .NET | how to track formula evaluation performance with Aspose.Cells | detect circular references using CalculationMonitor in Aspose.Cells | C# Aspose.Cells CalculationOptions with custom monitor
// Developer Intent: Implement a custom CalculationMonitor to capture detailed evaluation data for every cell while calculating formulas.
// Use Cases: Debug complex formula chains by reviewing the order and values of cell calculations. | Identify performance bottlenecks in large worksheets by extending the monitor to measure execution time per cell. | Monitor and report circular references without aborting the calculation process.
// AI Prompts: Generate a CalculationMonitor that timestamps each cell calculation and writes the log to a CSV file. | Provide code to aggregate total calculation time per worksheet and display a summary after workbook.CalculateFormula. | Create a monitor that skips cells in a user‑defined range but continues logging all other evaluations.

using System;
using Aspose.Cells;

// Shows how to create a workbook, insert values and formulas, and attach a custom LoggingMonitor (derived from AbstractCalculationMonitor) via CalculationOptions. The monitor records before/after events and circular‑reference alerts during workbook.CalculateFormula, then saves the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some cells with values and formulas
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["A2"].PutValue(10);
        sheet.Cells["A3"].Formula = "=A1+A2";          // Simple addition
        sheet.Cells["B1"].Formula = "=SUM(A1:A2)";    // Built‑in SUM function

        // Set up calculation options with a custom monitor to log each cell evaluation
        CalculationOptions options = new CalculationOptions
        {
            CalculationMonitor = new LoggingMonitor()
        };

        // Perform calculation using the options
        workbook.CalculateFormula(options);

        // Output final results (optional)
        Console.WriteLine($"A3 = {sheet.Cells["A3"].Value}");
        Console.WriteLine($"B1 = {sheet.Cells["B1"].Value}");

        // Save the workbook
        workbook.Save("LoggedCalculation.xlsx");
    }

    // Custom monitor that logs before and after each cell calculation
    class LoggingMonitor : AbstractCalculationMonitor
    {
        public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            Console.WriteLine($"Before calculating Sheet {sheetIndex}, Row {rowIndex}, Column {columnIndex}");
        }

        public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            Console.WriteLine($"After calculating Sheet {sheetIndex}, Row {rowIndex}, Column {columnIndex}");
            Console.WriteLine($"Original Value: {OriginalValue}, Calculated Value: {CalculatedValue}, Value Changed: {ValueChanged}");
        }

        public override bool OnCircular(System.Collections.IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected.");
            // Continue calculation despite circular reference
            return true;
        }
    }
}
