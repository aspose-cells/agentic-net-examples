using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsCalculationMonitorDemo
{
    // Custom monitor to track formula calculation progress
    public class CustomCalculationMonitor : AbstractCalculationMonitor
    {
        // Called before a cell is calculated
        public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            Console.WriteLine($"Before calculate - Sheet:{sheetIndex}, Row:{rowIndex}, Column:{columnIndex}");
        }

        // Called after a cell is calculated
        public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            Console.WriteLine($"After calculate  - Sheet:{sheetIndex}, Row:{rowIndex}, Column:{columnIndex}");
            Console.WriteLine($"  Original Value : {OriginalValue}");
            Console.WriteLine($"  Calculated Value: {CalculatedValue}");
            Console.WriteLine($"  Value Changed   : {ValueChanged}");
        }

        // Called when a circular reference is detected
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected:");
            while (circularCellsData.MoveNext())
            {
                Console.WriteLine($"  {circularCellsData.Current}");
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

            // Add sample formulas
            sheet.Cells["A1"].Formula = "=1+2";
            sheet.Cells["A2"].Formula = "=A1*3";
            sheet.Cells["A3"].Formula = "=SUM(A1:A2)";

            // Set up calculation options with the custom monitor
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = new CustomCalculationMonitor()
            };

            // Recalculate all formulas using the monitor
            workbook.CalculateFormula(options);

            // Output the final values
            Console.WriteLine($"A1 = {sheet.Cells["A1"].Value}");
            Console.WriteLine($"A2 = {sheet.Cells["A2"].Value}");
            Console.WriteLine($"A3 = {sheet.Cells["A3"].Value}");

            // Save the workbook
            workbook.Save("CalculationMonitorResult.xlsx");
        }
    }
}