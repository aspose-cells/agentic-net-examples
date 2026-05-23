using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsFormulaMonitoring
{
    // Custom monitor that logs each step of formula evaluation
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
            Console.WriteLine($"    Original Value: {OriginalValue}");
            Console.WriteLine($"    Calculated Value: {CalculatedValue}");
            Console.WriteLine($"    Value Changed: {ValueChanged}");
        }

        // Called when a circular reference is detected
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected among the following cells:");
            while (circularCellsData.MoveNext())
            {
                // Use dynamic to avoid compile‑time binding to unavailable members
                var cell = circularCellsData.Current as dynamic;
                if (cell != null)
                {
                    try
                    {
                        Console.WriteLine($"    Sheet {cell.SheetIndex}, Row {cell.Row}, Column {cell.Column}");
                    }
                    catch
                    {
                        // Fallback if properties are not accessible
                        Console.WriteLine("    (cell details unavailable)");
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
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Set up sample formulas
                sheet.Cells["A1"].Formula = "=1+2";               // Simple arithmetic
                sheet.Cells["A2"].Formula = "=A1*3";              // Dependent on A1
                sheet.Cells["A3"].Formula = "=SUM(A1:A2)";        // Built‑in function

                // Introduce a circular reference for demonstration
                sheet.Cells["B1"].Formula = "=B2";
                sheet.Cells["B2"].Formula = "=B1";

                // Configure calculation options with the custom monitor
                CalculationOptions options = new CalculationOptions
                {
                    CalculationMonitor = new LoggingCalculationMonitor(),
                    IgnoreError = false,
                    Recursive = true
                };

                // Perform calculation with monitoring
                workbook.CalculateFormula(options);

                // Output final values for verification
                Console.WriteLine("\nFinal cell values:");
                Console.WriteLine($"A1 = {sheet.Cells["A1"].Value}");
                Console.WriteLine($"A2 = {sheet.Cells["A2"].Value}");
                Console.WriteLine($"A3 = {sheet.Cells["A3"].Value}");
                Console.WriteLine($"B1 = {sheet.Cells["B1"].Value}");
                Console.WriteLine($"B2 = {sheet.Cells["B2"].Value}");

                // Save the workbook (optional, demonstrates lifecycle rule usage)
                string outputPath = "FormulaMonitoringResult.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"\nWorkbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}