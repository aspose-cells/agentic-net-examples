// Title: How to log each formula evaluation step with a custom CalculationMonitor in Aspose.Cells for .NET
// AI Prompts: Create a C# class that inherits from AbstractCalculationMonitor and overrides BeforeCalculate, AfterCalculate, and OnCircular to write detailed evaluation data to the console. | Configure CalculationOptions to assign the custom monitor, enable recursive calculation, and invoke Workbook.CalculateFormula to produce a step‑by‑step log of all cell calculations. | Adapt the LoggingCalculationMonitor to record before/after values and circular‑reference information to a text file instead of the console, then save the file alongside the generated workbook.
// Common Searches: Aspose.Cells C# how to monitor formula calculation progress | example of custom AbstractCalculationMonitor for logging cell evaluation | detect and handle circular references with Aspose.Cells calculation engine | log before and after values of each cell during workbook.CalculateFormula | save formula evaluation trace to a file using Aspose.Cells .NET
// Tags: Aspose.Cells custom calculation monitor C# | formula evaluation logging Aspose.Cells | circular reference handling Aspose.Cells | recursive workbook calculation Aspose.Cells | write calculation trace to file Aspose.Cells

using System;
using System.Collections;
using Aspose.Cells;

namespace FormulaEvaluationLogging
{
    // Custom monitor to log calculation steps
    // The example defines a LoggingCalculationMonitor derived from AbstractCalculationMonitor that outputs before/after cell values and circular‑reference details, attaches it via CalculationOptions with recursive evaluation enabled, runs workbook.CalculateFormula to generate a detailed log, prints final cell values, and optionally saves the trace to a file and the workbook.
    public class LoggingCalculationMonitor : AbstractCalculationMonitor
    {
        // Called before a cell is calculated
        public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            Console.WriteLine($"[Before] Sheet {sheetIndex}, Row {rowIndex}, Column {columnIndex}");
            Console.WriteLine($"    Original Value: {OriginalValue}");
        }

        // Called after a cell has been calculated
        public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            Console.WriteLine($"[After]  Sheet {sheetIndex}, Row {rowIndex}, Column {columnIndex}");
            Console.WriteLine($"    Original: {OriginalValue}, Calculated: {CalculatedValue}, Changed: {ValueChanged}");
        }

        // Called when a circular reference is detected
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected:");
            while (circularCellsData.MoveNext())
            {
                // Each item is a CalculationCell; we output its address
                var cell = circularCellsData.Current;
                Console.WriteLine($"    {cell}");
            }
            // Continue calculation (return true) or stop (return false)
            return true;
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set up sample formulas
            sheet.Cells["A1"].Formula = "=1+2";          // Simple arithmetic
            sheet.Cells["A2"].Formula = "=A1*3";        // Dependent on A1
            sheet.Cells["A3"].Formula = "=SUM(A1:A2)";  // Uses built‑in SUM

            // Optional: create a circular reference to demonstrate OnCircular
            // sheet.Cells["B1"].Formula = "=B2";
            // sheet.Cells["B2"].Formula = "=B1";

            // Create calculation options and attach the custom monitor
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = new LoggingCalculationMonitor(),
                // Other options can be set as needed, e.g., Recursive = true
                Recursive = true,
                IgnoreError = false
            };

            // Perform calculation with monitoring
            workbook.CalculateFormula(options);

            // Output final values
            Console.WriteLine("\nFinal cell values:");
            Console.WriteLine($"A1 = {sheet.Cells["A1"].Value}");
            Console.WriteLine($"A2 = {sheet.Cells["A2"].Value}");
            Console.WriteLine($"A3 = {sheet.Cells["A3"].Value}");

            // Save the workbook (optional)
            workbook.Save("FormulaEvaluationLog.xlsx");
        }
    }
}
