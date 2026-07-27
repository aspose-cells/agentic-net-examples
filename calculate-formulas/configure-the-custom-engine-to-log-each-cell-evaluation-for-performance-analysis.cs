using System;
using Aspose.Cells;

namespace AsposeCellsCustomEngineLogging
{
    // Custom calculation engine that logs each cell evaluation
    public class LoggingEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Log the cell address and the function being evaluated
            // For built‑in functions the FunctionName contains the name (e.g., "SUM")
            // For plain expressions FunctionName may be empty; we still log the cell.
            string cellAddress = data.Cell.Name; // e.g., "A3"
            string functionName = string.IsNullOrEmpty(data.FunctionName) ? "[Expression]" : data.FunctionName;

            Console.WriteLine($"Evaluating cell {cellAddress}, Function: {functionName}");

            // No custom calculation is performed; let the default engine handle the actual computation.
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data and formulas
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["A2"].PutValue(10);
            sheet.Cells["A3"].Formula = "=SUM(A1:A2)";      // Built‑in function
            sheet.Cells["B1"].Formula = "=A1*A2";          // Simple expression
            sheet.Cells["C1"].Formula = "=NOW()";          // Volatile function

            // Configure calculation options to use the custom logging engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new LoggingEngine()
            };

            // Perform calculation with the custom engine
            workbook.CalculateFormula(options);

            // Save the workbook (the file will contain the calculated values)
            workbook.Save("LoggedCalculation.xlsx");
        }
    }
}