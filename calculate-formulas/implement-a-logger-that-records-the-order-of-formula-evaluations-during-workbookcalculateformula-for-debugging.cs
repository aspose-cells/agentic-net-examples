// Title: How to create a custom AbstractCalculationMonitor in Aspose.Cells for .NET to log the order of formula evaluations during Workbook.CalculateFormula
// AI Prompts: Implement a C# class that inherits from Aspose.Cells.AbstractCalculationMonitor and records the cell address before each calculation and the original and new values after each calculation. | Set up CalculationOptions to assign the custom monitor, then call Workbook.CalculateFormula so the evaluation sequence is captured automatically. | After the workbook is calculated, retrieve the logged entries from the monitor and output them to the console or write them to a log file.
// Common Searches: aspnet log formula evaluation order using Aspose.Cells AbstractCalculationMonitor | debug Workbook.CalculateFormula sequence Aspose.Cells .NET example | track cell calculation dependencies with custom calculation monitor in Aspose.Cells | capture before and after values of formulas during calculation in Aspose.Cells
// Tags: Aspose.Cells custom calculation monitor | formula evaluation logging .NET | track cell calculation order Aspose.Cells | Workbook.CalculateFormula monitoring | log before after cell values Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsFormulaEvaluationLogger
{
    // Custom monitor that records the order of formula evaluations
    // Demonstrates how to subclass AbstractCalculationMonitor to record before/after events for each cell, attach the monitor via CalculationOptions, run Workbook.CalculateFormula, and then retrieve and display the evaluation log along with final cell values.
    public class EvaluationLogger : AbstractCalculationMonitor
    {
        // Stores log entries
        private readonly List<string> _log = new List<string>();

        // Called before a cell is calculated
        public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            string cellName = CellsHelper.CellIndexToName(rowIndex, columnIndex);
            _log.Add($"Before: Sheet{sheetIndex} {cellName}");
        }

        // Called after a cell is calculated
        public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            string cellName = CellsHelper.CellIndexToName(rowIndex, columnIndex);
            _log.Add($"After: Sheet{sheetIndex} {cellName} | " +
                     $"Original: {OriginalValue}, New: {CalculatedValue}, Changed: {ValueChanged}");
        }

        // Exposes the collected log
        public IEnumerable<string> GetLog()
        {
            return _log;
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data and formulas
            cells["A1"].PutValue(5);                 // Constant value
            cells["A2"].Formula = "=A1*2";           // Depends on A1
            cells["A3"].Formula = "=A2+10";          // Depends on A2
            cells["B1"].Formula = "=SUM(A1:A3)";     // Depends on A1, A2, A3
            cells["C1"].Formula = "=NOW()";          // Volatile function

            // Set up calculation options with the custom logger
            EvaluationLogger logger = new EvaluationLogger();
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = logger,
                // Enable calculation chain to ensure proper dependency tracking (optional)
                // This can be omitted if not needed
                // Recursive = true,
                // IgnoreError = false
            };

            // Perform calculation with monitoring
            workbook.CalculateFormula(options);

            // Output the evaluation order
            Console.WriteLine("Formula evaluation order:");
            foreach (var entry in logger.GetLog())
            {
                Console.WriteLine(entry);
            }

            // Optionally, display final cell values
            Console.WriteLine("\nFinal cell values:");
            Console.WriteLine($"A1 = {cells["A1"].Value}");
            Console.WriteLine($"A2 = {cells["A2"].Value}");
            Console.WriteLine($"A3 = {cells["A3"].Value}");
            Console.WriteLine($"B1 = {cells["B1"].Value}");
            Console.WriteLine($"C1 = {cells["C1"].Value}");

            // Save the workbook (demonstrates usage of save rule)
            workbook.Save("FormulaEvaluationLog.xlsx");
        }
    }
}
