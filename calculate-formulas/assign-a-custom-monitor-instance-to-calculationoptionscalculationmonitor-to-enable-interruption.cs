using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsCustomCalculationMonitor
{
    // Custom monitor to track formula calculation progress
    public class SampleCalculationMonitor : AbstractCalculationMonitor
    {
        public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            Console.WriteLine($"Before calculation: Sheet{sheetIndex}, Row{rowIndex}, Col{columnIndex}");
        }

        public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            Console.WriteLine($"After calculation: Sheet{sheetIndex}, Row{rowIndex}, Col{columnIndex}");
            Console.WriteLine($"Original Value: {OriginalValue}, Calculated Value: {CalculatedValue}, Value Changed: {ValueChanged}");
        }

        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected.");
            // Return false to continue calculation despite circular reference
            return false;
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set sample formulas
            sheet.Cells["A1"].Formula = "=1+2";
            sheet.Cells["A2"].Formula = "=A1*3";
            sheet.Cells["A3"].Formula = "=SUM(A1:A2)";

            // Create calculation options and assign the custom monitor
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = new SampleCalculationMonitor()
            };

            // Perform calculation with monitoring
            workbook.CalculateFormula(options);

            // Output the results
            Console.WriteLine($"A1: {sheet.Cells["A1"].Value}");
            Console.WriteLine($"A2: {sheet.Cells["A2"].Value}");
            Console.WriteLine($"A3: {sheet.Cells["A3"].Value}");
        }
    }
}