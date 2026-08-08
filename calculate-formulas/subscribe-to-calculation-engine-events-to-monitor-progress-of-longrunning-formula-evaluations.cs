// Title: Use Aspose.Cells CalculationMonitor in C# to Track Formula Evaluation and Detect Circular References
// Description: Shows how to extend AbstractCalculationMonitor with a ProgressCalculationMonitor, attach it through CalculationOptions, and run Workbook.CalculateFormula to log before/after cell processing, report original vs. calculated values, and handle circular references in .NET.
// Keywords: Aspose.Cells | CalculationMonitor | AbstractCalculationMonitor | C# | formula evaluation | circular reference detection | .NET | Workbook.CalculateFormula | progress logging | long‑running calculations
// Common Searches: Aspose.Cells custom calculation monitor example | track formula calculation progress C# | detect circular references Aspose.Cells | log before after cell calculation Aspose.Cells | how to use CalculationOptions with monitor
// Developer Intent: Subscribe to calculation engine events to monitor the progress of long‑running formula evaluations.
// Use Cases: Console‑based debugging of large workbooks by logging each cell's calculation start and end. | Real‑time detection and reporting of circular references while allowing the calculation to continue. | Combining the monitor with options such as Recursive and IgnoreError to control evaluation flow and capture detailed metrics.
// AI Prompts: Create a CalculationMonitor that writes before/after details to a file instead of the console. | Add logic to count processed cells and display the total after Workbook.CalculateFormula finishes. | Show how to abort calculation inside OnCircular based on a custom threshold.

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsCalculationMonitorDemo
{
    // Custom monitor to track calculation progress
    // Shows how to extend AbstractCalculationMonitor with a ProgressCalculationMonitor, attach it through CalculationOptions, and run Workbook.CalculateFormula to log before/after cell processing, report original vs. calculated values, and handle circular references in .NET.
    public class ProgressCalculationMonitor : AbstractCalculationMonitor
    {
        // Called before each cell is calculated
        public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            Console.WriteLine($"[Before] Sheet {sheetIndex}, Row {rowIndex}, Column {columnIndex}");
        }

        // Called after each cell is calculated
        public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            Console.WriteLine($"[After]  Sheet {sheetIndex}, Row {rowIndex}, Column {columnIndex}");
            Console.WriteLine($"    Original: {OriginalValue}, New: {CalculatedValue}, Changed: {ValueChanged}");
        }

        // Called when a circular reference is detected
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected:");
            while (circularCellsData.MoveNext())
            {
                Console.WriteLine($"    {circularCellsData.Current}");
            }
            // Continue calculation for circular cells
            return true;
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells with data and formulas (simulate a long‑running scenario)
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].Formula = "=A1+A2";               // Simple sum
            sheet.Cells["B1"].Formula = "=A3*5";                // Dependent formula
            sheet.Cells["C1"].Formula = "=SUM(A1:A3)";          // Built‑in function
            sheet.Cells["D1"].Formula = "=IF(A1>5,\"High\",\"Low\")";

            // Create calculation options and attach the custom monitor
            CalculationOptions calcOptions = new CalculationOptions
            {
                CalculationMonitor = new ProgressCalculationMonitor(),
                Recursive = true,          // Ensure dependent cells are calculated
                IgnoreError = false       // Show errors if they occur
            };

            // Perform calculation with monitoring
            workbook.CalculateFormula(calcOptions);

            // Output final results
            Console.WriteLine("\nFinal cell values:");
            Console.WriteLine($"A3 = {sheet.Cells["A3"].Value}");
            Console.WriteLine($"B1 = {sheet.Cells["B1"].Value}");
            Console.WriteLine($"C1 = {sheet.Cells["C1"].Value}");
            Console.WriteLine($"D1 = {sheet.Cells["D1"].Value}");

            // Save the workbook (demonstrates lifecycle rule usage)
            workbook.Save("CalculationMonitorResult.xlsx");
        }
    }
}
