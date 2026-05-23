using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Custom monitor to track calculation progress
    public class SampleCalculationMonitor : AbstractCalculationMonitor
    {
        // Called before each cell is calculated
        public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            Console.WriteLine($"Before calculation: Sheet{sheetIndex}, Row{rowIndex}, Col{columnIndex}");
        }

        // Called after each cell is calculated
        public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            Console.WriteLine($"After calculation: Sheet{sheetIndex}, Row{rowIndex}, Col{columnIndex}");
            Console.WriteLine($"Original Value: {OriginalValue}, Calculated Value: {CalculatedValue}, Value Changed: {ValueChanged}");
        }

        // Called when a circular reference is detected
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected.");
            // Continue calculation for circular cells
            return true;
        }
    }

    public class CalculationWithMonitorDemo
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data and formulas
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["A2"].Formula = "=A1*2";
            sheet.Cells["A3"].Formula = "=A2+10";

            // Set up calculation options with the custom monitor
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = new SampleCalculationMonitor(),
                // Optional: keep default behavior for other options
                IgnoreError = false,
                Recursive = true
            };

            // Recalculate all formulas using the options
            workbook.CalculateFormula(options);

            // Output the final values
            Console.WriteLine($"A1 = {sheet.Cells["A1"].Value}");
            Console.WriteLine($"A2 = {sheet.Cells["A2"].Value}");
            Console.WriteLine($"A3 = {sheet.Cells["A3"].Value}");

            // Save the workbook (using the standard save rule)
            workbook.Save("CalculationWithMonitorDemo.xlsx");
        }
    }

    // Entry point
    class Program
    {
        static void Main()
        {
            CalculationWithMonitorDemo.Run();
        }
    }
}