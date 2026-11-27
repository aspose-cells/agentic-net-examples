// Title: Create a custom calculation engine in Aspose.Cells for .NET to log the order of formula evaluations during Workbook.CalculateFormula
// Common Searches: how to capture formula evaluation sequence with Aspose.Cells CalculateFormula in C# | using AbstractCalculationEngine to trace Excel formula execution in Aspose.Cells | debugging cross-sheet formula calculations by logging function calls in Aspose.Cells .NET | recording each function name and cell reference during workbook calculation with Aspose.Cells | custom logging engine for Aspose.Cells to monitor SUM and AVERAGE evaluations
// Tags: custom AbstractCalculationEngine logging Aspose.Cells | formula evaluation order tracking .NET | Workbook.CalculateFormula debug logging | cross-sheet calculation tracing Aspose.Cells | minimal SUM AVERAGE implementation custom engine

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace FormulaEvaluationLogger
{
    // Custom calculation engine that logs each function evaluation order.
    // The example defines a LoggingEngine class that inherits from AbstractCalculationEngine and records every evaluated function together with its cell reference in an EvaluationLog list. It provides simple implementations for SUM and AVERAGE while delegating other functions to Aspose.Cells' built‑in engine. A workbook with inter‑dependent and cross‑sheet formulas is built, CalculationOptions are set with the custom engine and Recursive = true, and Workbook.CalculateFormula is executed. After calculation the program prints the formula results and outputs the captured evaluation order, enabling developers to debug and understand the sequence of formula calculations.
    public class LoggingEngine : AbstractCalculationEngine
    {
        // Stores the evaluation log entries.
        public List<string> EvaluationLog { get; } = new List<string>();

        // This method is called by Aspose.Cells for every function encountered during calculation.
        public override void Calculate(CalculationData data)
        {
            // Build a readable identifier for the cell being evaluated.
            string cellRef;
            if (data.Cell != null && !string.IsNullOrEmpty(data.Cell.Name))
            {
                cellRef = data.Cell.Name;
            }
            else if (data.Cell != null)
            {
                cellRef = $"R{data.Cell.Row + 1}C{data.Cell.Column + 1}";
            }
            else
            {
                cellRef = "UnknownCell";
            }

            // Record the function name together with the cell reference.
            string logEntry = $"{cellRef} -> {data.FunctionName}";
            EvaluationLog.Add(logEntry);

            // Minimal implementation of a few common functions so that the workbook can still be calculated while we are logging.
            if (string.Equals(data.FunctionName, "SUM", StringComparison.OrdinalIgnoreCase))
            {
                double sum = 0;
                for (int i = 0; i < data.ParamCount; i++)
                {
                    object val = data.GetParamValue(i);
                    if (val is double d) sum += d;
                    else if (val is int iVal) sum += iVal;
                    else if (val is string s && double.TryParse(s, out double parsed)) sum += parsed;
                }
                data.CalculatedValue = sum;
            }
            else if (string.Equals(data.FunctionName, "AVERAGE", StringComparison.OrdinalIgnoreCase))
            {
                double sum = 0;
                int count = 0;
                for (int i = 0; i < data.ParamCount; i++)
                {
                    object val = data.GetParamValue(i);
                    if (val is double d) { sum += d; count++; }
                    else if (val is int iVal) { sum += iVal; count++; }
                    else if (val is string s && double.TryParse(s, out double parsed)) { sum += parsed; count++; }
                }
                data.CalculatedValue = count > 0 ? sum / count : 0.0;
            }
            // For any other function we do not set CalculatedValue;
            // Aspose.Cells will then use its built‑in engine to compute the result.
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // ------------------------------------------------------------
                // 1. Build a sample workbook with several inter‑dependent formulas.
                // ------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Simple numeric data.
                cells["A1"].PutValue(10);
                cells["A2"].PutValue(20);
                cells["A3"].PutValue(30);

                // Formulas that depend on the above data.
                cells["B1"].Formula = "=SUM(A1:A3)";          // Expected 60
                cells["B2"].Formula = "=AVERAGE(A1:A3)";      // Expected 20
                cells["B3"].Formula = "=B1*2";                // Depends on B1

                // A formula that references another sheet (to test recursive option).
                int secondIndex = workbook.Worksheets.Add();
                Worksheet sheet2 = workbook.Worksheets[secondIndex];
                sheet2.Name = "Second";
                sheet2.Cells["A1"].PutValue(5);
                sheet2.Cells["A2"].Formula = "=First!B1 + A1"; // First!B1 is 60, plus 5 => 65

                // ------------------------------------------------------------
                // 2. Prepare calculation options with our custom logging engine.
                // ------------------------------------------------------------
                CalculationOptions calcOptions = new CalculationOptions
                {
                    CustomEngine = new LoggingEngine(),
                    Recursive = true   // Ensure cross‑sheet dependencies are calculated.
                };

                // ------------------------------------------------------------
                // 3. Perform full workbook calculation.
                // ------------------------------------------------------------
                workbook.CalculateFormula(calcOptions);

                // ------------------------------------------------------------
                // 4. Output the results of the formulas.
                // ------------------------------------------------------------
                Console.WriteLine("=== Formula Results ===");
                Console.WriteLine($"B1 (SUM): {cells["B1"].StringValue}");
                Console.WriteLine($"B2 (AVERAGE): {cells["B2"].StringValue}");
                Console.WriteLine($"B3 (B1*2): {cells["B3"].StringValue}");
                Console.WriteLine($"Second!A2 (cross‑sheet): {sheet2.Cells["A2"].StringValue}");

                // ------------------------------------------------------------
                // 5. Dump the evaluation order captured by the logger.
                // ------------------------------------------------------------
                Console.WriteLine("\n=== Formula Evaluation Order ===");
                foreach (string entry in ((LoggingEngine)calcOptions.CustomEngine).EvaluationLog)
                {
                    Console.WriteLine(entry);
                }

                // Optional: Save the workbook to verify that calculations succeeded.
                // string outPath = "LoggedFormulaEvaluation.xlsx";
                // workbook.Save(outPath, SaveFormat.Xlsx);
                // Console.WriteLine($"Workbook saved to {Path.GetFullPath(outPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
