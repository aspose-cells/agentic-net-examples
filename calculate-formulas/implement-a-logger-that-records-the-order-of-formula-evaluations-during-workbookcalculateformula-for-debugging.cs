using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace FormulaEvaluationLogger
{
    // Custom monitor that records the order of cell calculations
    public class EvaluationLogger : AbstractCalculationMonitor
    {
        private readonly List<string> _log;

        public EvaluationLogger(List<string> log)
        {
            _log = log;
        }

        public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            // Record before calculation
            _log.Add($"Before: Sheet{sheetIndex}, Row{rowIndex}, Col{columnIndex}");
        }

        public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            // Record after calculation
            _log.Add($"After: Sheet{sheetIndex}, Row{rowIndex}, Col{columnIndex}");
        }

        public override bool OnCircular(System.Collections.IEnumerator circularCellsData)
        {
            // No special handling for circular references in this demo
            return base.OnCircular(circularCellsData);
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
            cells["A1"].PutValue(5);
            cells["A2"].Formula = "=A1*2";
            cells["A3"].Formula = "=A2+10";
            cells["B1"].Formula = "=SUM(A1:A3)";

            // Prepare a log container
            List<string> evaluationLog = new List<string>();

            // Set up calculation options with the custom monitor
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = new EvaluationLogger(evaluationLog)
            };

            // Perform calculation with monitoring
            workbook.CalculateFormula(options);

            // Output the recorded evaluation order
            Console.WriteLine("Formula evaluation order:");
            foreach (string entry in evaluationLog)
            {
                Console.WriteLine(entry);
            }

            // Optionally display final values
            Console.WriteLine("\nFinal cell values:");
            Console.WriteLine($"A1 = {cells["A1"].Value}");
            Console.WriteLine($"A2 = {cells["A2"].Value}");
            Console.WriteLine($"A3 = {cells["A3"].Value}");
            Console.WriteLine($"B1 = {cells["B1"].Value}");

            // Save the workbook (using the standard save method)
            workbook.Save("FormulaEvaluationLog.xlsx");
        }
    }
}