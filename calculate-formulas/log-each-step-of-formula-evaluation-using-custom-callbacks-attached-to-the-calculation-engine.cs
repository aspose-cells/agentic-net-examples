// Title: Log Formula Evaluation Steps with a Custom Calculation Monitor in Aspose.Cells (C#)
// Description: This example shows how to create a LoggingCalculationMonitor that inherits from AbstractCalculationMonitor and overrides BeforeCalculate, AfterCalculate, and OnCircular. The monitor writes cell coordinates, original values, calculated results, and circular‑reference details to the console. It is attached to CalculationOptions, which are then passed to Workbook.CalculateFormula to execute the workbook while providing a step‑by‑step trace, followed by output of final values and optional workbook saving.
// Keywords: Aspose.Cells | C# | CalculationMonitor | AbstractCalculationMonitor | formula logging | custom callbacks | cell calculation trace | circular reference detection | CalculationOptions | Workbook.CalculateFormula | debug Excel formulas | spreadsheet performance monitoring
// Common Searches: Aspose.Cells log each formula calculation | custom calculation monitor C# example | how to detect circular references with Aspose.Cells | trace formula evaluation steps Aspose.Cells | before and after calculate callbacks Aspose.Cells
// Developer Intent: Implement a custom monitor to capture detailed information before and after each cell’s formula is evaluated and to handle circular references during workbook calculation.
// Use Cases: Debug complex workbooks by printing the original and computed values of every formula cell. | Automatically identify circular references and decide whether to continue or abort the calculation. | Integrate logging with existing CalculationOptions to audit spreadsheet calculations in production or testing environments.
// AI Prompts: Write a C# class that extends AbstractCalculationMonitor and logs cell coordinates, original value, and calculated result for every formula evaluation in Aspose.Cells. | Demonstrate how to assign a custom CalculationMonitor to CalculationOptions and invoke Workbook.CalculateFormula with logging enabled. | Provide code that captures circular reference information inside OnCircular, prints the involved cells, and returns true to allow the engine to continue.

using System;
using System.Collections;
using Aspose.Cells;

namespace FormulaEvaluationLoggingDemo
{
    // Custom monitor to log each step of formula calculation
    // This example shows how to create a LoggingCalculationMonitor that inherits from AbstractCalculationMonitor and overrides BeforeCalculate, AfterCalculate, and OnCircular. The monitor writes cell coordinates, original values, calculated results, and circular‑reference details to the console. It is attached to CalculationOptions, which are then passed to Workbook.CalculateFormula to execute the workbook while providing a step‑by‑step trace, followed by output of final values and optional workbook saving.
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
        }

        // Called when a circular reference is detected
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected among the following cells:");
            while (circularCellsData.MoveNext())
            {
                var cell = circularCellsData.Current;
                if (cell != null)
                {
                    // Try to retrieve SheetIndex, Row, Column via reflection (API may vary by version)
                    var type = cell.GetType();
                    var sheetIdxProp = type.GetProperty("SheetIndex");
                    var rowProp = type.GetProperty("Row");
                    var colProp = type.GetProperty("Column");

                    if (sheetIdxProp != null && rowProp != null && colProp != null)
                    {
                        int sheetIdx = (int)sheetIdxProp.GetValue(cell);
                        int row = (int)rowProp.GetValue(cell);
                        int col = (int)colProp.GetValue(cell);
                        Console.WriteLine($"    Sheet {sheetIdx}, Row {row}, Column {col}");
                    }
                    else
                    {
                        Console.WriteLine($"    Cell info: {cell}");
                    }
                }
            }
            // Continue calculation for circular cells
            return true;
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate cells with sample data and formulas
                sheet.Cells["A1"].PutValue(5);
                sheet.Cells["A2"].PutValue(10);
                sheet.Cells["A3"].Formula = "=A1+A2";          // Simple addition
                sheet.Cells["A4"].Formula = "=A3*2";           // Dependent on A3
                sheet.Cells["A5"].Formula = "=SUM(A1:A4)";     // Aggregate function

                // Introduce a circular reference for demonstration
                sheet.Cells["B1"].Formula = "=B2";
                sheet.Cells["B2"].Formula = "=B1";

                // Set up calculation options with the custom monitor
                CalculationOptions options = new CalculationOptions
                {
                    CalculationMonitor = new LoggingCalculationMonitor(),
                    // Optional: keep default behavior for other options
                    IgnoreError = false,
                    Recursive = true
                };

                // Perform calculation with monitoring
                workbook.CalculateFormula(options);

                // Output final values of the calculated cells
                Console.WriteLine("\nFinal calculated values:");
                Console.WriteLine($"A3 = {sheet.Cells["A3"].Value}");
                Console.WriteLine($"A4 = {sheet.Cells["A4"].Value}");
                Console.WriteLine($"A5 = {sheet.Cells["A5"].Value}");

                // Save the workbook (optional)
                string outputPath = "FormulaEvaluationLoggingDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"\nWorkbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
