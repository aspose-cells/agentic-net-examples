// Title: Log every formula cell evaluation with a custom Aspose.Cells calculation engine in C#
// AI Prompts: Create a C# class that inherits AbstractCalculationEngine and writes the worksheet name, cell address, and function name to the console inside the Calculate method. | Configure CalculationOptions to assign the custom LoggingEngine, then call Workbook.CalculateFormula to trigger logging for all formula cells. | Enhance the LoggingEngine to measure and output the elapsed time for each formula evaluation together with the cell details.
// Common Searches: aspnet how to use a custom calculation engine in Aspose.Cells to log formula evaluations | c# Aspose.Cells log each cell formula execution for performance monitoring | example of overriding AbstractCalculationEngine to capture formula name and cell location in Aspose.Cells | measure formula calculation time per cell with Aspose.Cells custom engine C#
// Tags: custom calculation engine Aspose.Cells | log formula evaluation per cell C# | Aspose.Cells calculationoptions customengine | performance monitoring spreadsheet formulas | console logging cell calculations Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsCustomEngineLogging
{
    // Custom calculation engine that logs each function evaluation (i.e., each cell with a formula)
    // The example defines a LoggingEngine class derived from AbstractCalculationEngine that overrides ProcessBuiltInFunctions and Calculate to write worksheet name, row, column, and function name to the console for each formula cell. CalculationOptions is set to use this engine, wb.CalculateFormula triggers the logging, and the workbook is saved after verification.
    public class LoggingEngine : AbstractCalculationEngine
    {
        // Enable processing of built‑in functions so that the engine is invoked for every formula.
        public override bool ProcessBuiltInFunctions => true;

        // Called for each function (built‑in or custom) during calculation.
        public override void Calculate(CalculationData data)
        {
            // Log basic information about the cell being evaluated.
            Console.WriteLine($"[Engine] Sheet: {data.Worksheet.Name}, Row: {data.CellRow}, Column: {data.CellColumn}, Function: {data.FunctionName}");

            // Do not modify data.CalculatedValue here – let the default engine continue the calculation.
            // By not setting CalculatedValue, Aspose.Cells will compute the function normally.
        }

        // No special force‑recalculation logic required.
        public override bool ForceRecalculate(string functionName) => false;
    }

    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate some data and formulas
            ws.Cells["A1"].PutValue(10);
            ws.Cells["A2"].PutValue(20);
            ws.Cells["A3"].Formula = "=SUM(A1:A2)";          // Built‑in function
            ws.Cells["B1"].Formula = "=A1*A2";              // Arithmetic expression
            ws.Cells["B2"].Formula = "=HYPERLINK(\"http://example.com\",\"Link\")"; // Built‑in function with string

            // ---------- Configure calculation options with the custom logging engine ----------
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new LoggingEngine(),
                // Other options can be set as needed, e.g., IgnoreError = false, Recursive = true
            };

            // Perform calculation – the LoggingEngine will be invoked for each formula cell.
            wb.CalculateFormula(options);

            // Output results to verify that calculations succeeded.
            Console.WriteLine($"A3 (SUM) = {ws.Cells["A3"].Value}");
            Console.WriteLine($"B1 (A1*A2) = {ws.Cells["B1"].Value}");
            Console.WriteLine($"B2 (HYPERLINK) = {ws.Cells["B2"].StringValue}");

            // ---------- Save the workbook ----------
            wb.Save("CustomEngineLoggingDemo.xlsx");
            wb.Save("CustomEngineLoggingDemo.pdf");
        }
    }
}
