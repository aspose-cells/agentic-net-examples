// Title: C# – Custom AbstractCalculationMonitor to Log Cells Before and After Formula Evaluation in Aspose.Cells
// Description: Demonstrates how to derive from AbstractCalculationMonitor, override BeforeCalculate (and optionally AfterCalculate and OnCircular) to log sheet, row, column, original and calculated values during Workbook.CalculateFormula in Aspose.Cells for .NET.
// Keywords: Aspose.Cells | AbstractCalculationMonitor | BeforeCalculate | AfterCalculate | CalculationMonitor | C# | .NET | formula calculation | debugging formulas | circular reference detection | CalculationOptions | Workbook.CalculateFormula | cell inspection | logging
// Common Searches: custom calculation monitor Aspose.Cells C# | log cell values before formula calculation Aspose.Cells | override BeforeCalculate in Aspose.Cells | detect circular references with Aspose.Cells monitor | set CalculationOptions.CalculationMonitor .NET
// Developer Intent: Create a subclass of AbstractCalculationMonitor that records cell coordinates and values during each step of formula calculation.
// Use Cases: Debug complex workbooks by printing original and resulting values for every evaluated cell. | Identify and report circular references while the calculation engine processes formulas. | Produce an audit trail of formula evaluation for compliance or performance analysis.
// AI Prompts: Generate a C# version of the monitor that writes inspection logs to a file instead of the console. | Show how to skip calculation for cells containing a specific comment or custom tag using AbstractCalculationMonitor. | Explain how to integrate the custom calculation monitor with parallel workbook calculations in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsCalculationMonitorDemo
{
    // Custom monitor that inspects each cell before it is calculated
    // Demonstrates how to derive from AbstractCalculationMonitor, override BeforeCalculate (and optionally AfterCalculate and OnCircular) to log sheet, row, column, original and calculated values during Workbook.CalculateFormula in Aspose.Cells for .NET.
    public class CellInspectionMonitor : AbstractCalculationMonitor
    {
        // This method is called by the calculation engine before a cell is evaluated
        public override void BeforeCalculate(int sheetIndex, int rowIndex, int colIndex)
        {
            // Retrieve the workbook and worksheet via the calculation context is not directly available here,
            // but we can log the cell coordinates and its original value.
            Console.WriteLine($"[Before] Sheet: {sheetIndex}, Row: {rowIndex}, Column: {colIndex}");

            // OriginalValue holds the value before calculation; useful for debugging.
            if (OriginalValue != null)
            {
                Console.WriteLine($"    Original Value: {OriginalValue}");
            }
            else
            {
                Console.WriteLine("    Original Value: <null>");
            }
        }

        // Optional: also show after calculation for completeness
        public override void AfterCalculate(int sheetIndex, int rowIndex, int colIndex)
        {
            Console.WriteLine($"[After] Sheet: {sheetIndex}, Row: {rowIndex}, Column: {colIndex}");
            Console.WriteLine($"    Calculated Value: {CalculatedValue}, Value Changed: {ValueChanged}");
        }

        // Optional: handle circular references
        public override bool OnCircular(System.Collections.IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected during calculation.");
            while (circularCellsData.MoveNext())
            {
                Console.WriteLine($"    Circular Cell: {circularCellsData.Current}");
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

            // Populate some sample data
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].Formula = "=A1+A2";          // Simple addition
            sheet.Cells["B1"].Formula = "=SUM(A1:A3)";    // Sum including a formula cell
            sheet.Cells["C1"].Formula = "=NOW()";        // Volatile function

            // Instantiate the custom calculation monitor
            CellInspectionMonitor monitor = new CellInspectionMonitor();

            // Configure calculation options to use the monitor
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = monitor,
                IgnoreError = false,
                Recursive = true
            };

            // Perform formula calculation with monitoring
            workbook.CalculateFormula(options);

            // Output final values for verification
            Console.WriteLine("\nFinal cell values:");
            Console.WriteLine($"A3 = {sheet.Cells["A3"].Value}");
            Console.WriteLine($"B1 = {sheet.Cells["B1"].Value}");
            Console.WriteLine($"C1 = {sheet.Cells["C1"].Value}");

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("CellInspectionDemo.xlsx");
        }
    }
}
