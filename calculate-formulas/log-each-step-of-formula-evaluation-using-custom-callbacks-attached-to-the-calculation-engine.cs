// Title: C# – Log Formula Evaluation Steps with a Custom CalculationMonitor in Aspose.Cells
// Description: Demonstrates how to attach a custom CalculationMonitor to Aspose.Cells' calculation engine to log each formula evaluation step. The LoggingCalculationMonitor overrides BeforeCalculate, AfterCalculate, and OnCircular to output original, calculated, and change information, handle division‑by‑zero errors, and report circular references. The sample creates a workbook, adds values and formulas, configures CalculationOptions with the monitor, and runs workbook.CalculateFormula while printing detailed logs to the console.
// Keywords: Aspose.Cells | C# | CalculationMonitor | formula logging | custom callbacks | BeforeCalculate | AfterCalculate | circular reference detection | division by zero handling | debug workbook formulas | Workbook.CalculateFormula
// Common Searches: Aspose.Cells log each formula calculation step | How to use CalculationMonitor in Aspose.Cells .NET | Detect circular references with Aspose.Cells | Track original and calculated cell values Aspose.Cells | Debug formula errors using custom monitor Aspose
// Developer Intent: The developer wants to monitor and record every stage of formula calculation—including pre‑ and post‑values and circular reference events—by implementing a custom CalculationMonitor in Aspose.Cells for .NET.
// Use Cases: Debug complex spreadsheets by printing before/after values for each calculated cell. | Identify and log circular reference chains to troubleshoot formula logic. | Capture and report calculation errors such as division by zero while preserving step‑by‑step logs.
// AI Prompts: Create a LoggingCalculationMonitor that adds timestamps and writes logs to a file instead of the console. | Show how to redirect the monitor's output to a structured JSON log for later analysis. | Modify the monitor to skip logging for cells without formulas and only log formula cells.

using System;
using System.Collections;
using Aspose.Cells;

namespace FormulaEvaluationLoggingDemo
{
    // Custom monitor to log each step of formula calculation
    // Demonstrates how to attach a custom CalculationMonitor to Aspose.Cells' calculation engine to log each formula evaluation step. The LoggingCalculationMonitor overrides BeforeCalculate, AfterCalculate, and OnCircular to output original, calculated, and change information, handle division‑by‑zero errors, and report circular references. The sample creates a workbook, adds values and formulas, configures CalculationOptions with the monitor, and runs workbook.CalculateFormula while printing detailed logs to the console.
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
            Console.WriteLine(new string('-', 40));
        }

        // Called when a circular reference is detected
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected!");
            while (circularCellsData.MoveNext())
            {
                Console.WriteLine($"    Circular cell: {circularCellsData.Current}");
            }
            // Continue calculation for circular cells
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

            // Populate some sample data and formulas
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].Formula = "=A1+A2";          // Simple addition
            sheet.Cells["A4"].Formula = "=A3*2";           // Dependent on A3
            sheet.Cells["B1"].Formula = "=SUM(A1:A4)";     // Aggregate function
            sheet.Cells["B2"].Formula = "=B1/0";           // Intentional error to show handling

            // Set up calculation options with the custom monitor
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = new LoggingCalculationMonitor(),
                IgnoreError = false,   // Do not ignore errors so they are reported
                Recursive = true
            };

            // Perform calculation with monitoring
            workbook.CalculateFormula(options);

            // Output final values for verification
            Console.WriteLine("\nFinal cell values:");
            Console.WriteLine($"A3 = {sheet.Cells["A3"].Value}");
            Console.WriteLine($"A4 = {sheet.Cells["A4"].Value}");
            Console.WriteLine($"B1 = {sheet.Cells["B1"].Value}");
            Console.WriteLine($"B2 = {sheet.Cells["B2"].Value}");
        }
    }
}
