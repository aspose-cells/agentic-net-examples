using System;
using Aspose.Cells;

namespace AsposeCellsCustomEngineLogging
{
    // Custom calculation engine that logs each cell evaluation.
    public class LoggingEngine : AbstractCalculationEngine
    {
        // No need to process built‑in functions; we only log and let the default engine handle the calculation.
        public override bool ProcessBuiltInFunctions => false;

        public override void Calculate(CalculationData data)
        {
            // Log details of the cell being evaluated.
            Console.WriteLine(
                $"Evaluating cell R{data.CellRow + 1}C{data.CellColumn + 1} " +
                $"(Sheet: \"{data.Worksheet.Name}\") " +
                $"Function: {data.FunctionName}");

            // Do not set CalculatedValue here; the default engine will compute the result.
        }

        // No special force‑recalculation logic required.
        public override bool ForceRecalculate(string functionName) => false;
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and obtain the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data.
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].Formula = "=A1+A2";          // Simple addition
            sheet.Cells["B1"].Formula = "=SUM(A1:A2)";    // Built‑in SUM
            sheet.Cells["B2"].Formula = "=NOW()";        // Volatile function

            // Configure calculation options to use the custom logging engine.
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new LoggingEngine(),
                // Other options can be set as needed.
                IgnoreError = false,
                Recursive = true
            };

            // Perform formula calculation with the custom engine.
            workbook.CalculateFormula(options);

            // Output the calculated values to verify correctness.
            Console.WriteLine($"A3 = {sheet.Cells["A3"].Value}");
            Console.WriteLine($"B1 = {sheet.Cells["B1"].Value}");
            Console.WriteLine($"B2 = {sheet.Cells["B2"].Value}");

            // Save the workbook (lifecycle rule: use provided save method).
            workbook.Save("LoggingEngineDemo.xlsx");
        }
    }
}