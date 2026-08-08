// Title: Log Formula Evaluation Order with a Custom AbstractCalculationMonitor in Aspose.Cells for .NET
// Description: Demonstrates how to subclass AbstractCalculationMonitor to capture before‑ and after‑calculate events for each cell during Workbook.CalculateFormula, storing sheet, row, column, original and calculated values for debugging complex formula dependencies.
// Keywords: Aspose.Cells | .NET | AbstractCalculationMonitor | CalculationMonitor | formula evaluation logging | Workbook.CalculateFormula | debug formulas | custom calculation monitor | cell calculation events | evaluation order
// Common Searches: Aspose.Cells log formula evaluation order | How to use AbstractCalculationMonitor in C# | Record before and after cell values during calculation | Debug formula dependencies Aspose.Cells | Custom calculation monitor example .NET
// Developer Intent: Create a calculation monitor that records each cell's before‑ and after‑evaluation details during Workbook.CalculateFormula for debugging purposes.
// Use Cases: Trace the exact sequence in which formulas are calculated to diagnose dependency issues. | Identify volatile functions (e.g., NOW, TODAY) and see when they are evaluated. | Compare logged evaluation order with expected precedence to verify calculation chain correctness. | Generate a change‑log of cells whose values were altered during a calculation run.
// AI Prompts: Write a C# class that extends AbstractCalculationMonitor, logs sheet, row, column, original and calculated values before and after each cell calculation, and attach it to CalculationOptions for Workbook.CalculateFormula. | Show how to filter the EvaluationLogger output to list only cells whose values changed during the calculation. | Explain how to enable the calculation chain in Aspose.Cells settings and why it matters when using a custom calculation monitor.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace FormulaEvaluationLogger
{
    // Custom monitor that records the order of formula evaluations
    // Demonstrates how to subclass AbstractCalculationMonitor to capture before‑ and after‑calculate events for each cell during Workbook.CalculateFormula, storing sheet, row, column, original and calculated values for debugging complex formula dependencies.
    public class EvaluationLogger : AbstractCalculationMonitor
    {
        // Stores log entries in the order they occur
        private readonly List<string> _log = new List<string>();

        // Called before a cell is calculated
        public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            _log.Add($"Before: Sheet{sheetIndex}, Row{rowIndex}, Col{columnIndex}");
        }

        // Called after a cell is calculated
        public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            // Use properties from AbstractCalculationMonitor to get details
            string entry = $"After: Sheet{sheetIndex}, Row{rowIndex}, Col{columnIndex}, " +
                           $"Original={OriginalValue}, Calculated={CalculatedValue}, Changed={ValueChanged}";
            _log.Add(entry);
        }

        // Expose the collected log
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

            // Set up sample formulas with dependencies
            sheet.Cells["A1"].Formula = "=1+2";          // Simple formula
            sheet.Cells["A2"].Formula = "=A1*3";        // Depends on A1
            sheet.Cells["A3"].Formula = "=SUM(A1:A2)";  // Depends on A1 and A2
            sheet.Cells["B1"].Formula = "=NOW()";      // Volatile function

            // Create the custom calculation monitor
            EvaluationLogger logger = new EvaluationLogger();

            // Configure calculation options to use the monitor
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = logger,
                // Enable calculation chain to ensure proper dependency tracking (optional)
                // This can be set via workbook settings if needed:
                // workbook.Settings.FormulaSettings.EnableCalculationChain = true;
            };

            // Perform formula calculation with monitoring
            workbook.CalculateFormula(options);

            // Output the evaluation order
            Console.WriteLine("Formula Evaluation Log:");
            foreach (string entry in logger.GetLog())
            {
                Console.WriteLine(entry);
            }

            // Save the workbook (optional, demonstrates lifecycle compliance)
            workbook.Save("FormulaEvaluationLog.xlsx");
        }
    }
}
