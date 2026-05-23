using System;
using Aspose.Cells;

namespace AsposeCellsCalculationMonitorDemo
{
    // Custom monitor that inspects each cell before it is calculated
    public class CustomCalculationMonitor : AbstractCalculationMonitor
    {
        // This method is called by the calculation engine before a cell is evaluated
        public override void BeforeCalculate(int sheetIndex, int rowIndex, int colIndex)
        {
            // Output the location of the cell that is about to be calculated
            Console.WriteLine($"Before calculating cell - Sheet: {sheetIndex}, Row: {rowIndex}, Column: {colIndex}");
        }

        // Optional: you can also override AfterCalculate if you need post‑calculation info
        public override void AfterCalculate(int sheetIndex, int rowIndex, int colIndex)
        {
            // Example of using the monitor's properties after calculation
            if (ValueChanged)
            {
                Console.WriteLine($"Cell changed from [{OriginalValue}] to [{CalculatedValue}]");
            }
        }

        // Optional: handle circular references
        public override bool OnCircular(System.Collections.IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected.");
            // Continue calculation for circular cells
            return true;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle rule: use provided creation logic)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data and formulas
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].Formula = "=A1+A2";   // This cell will be inspected by the monitor

            // Set up calculation options and attach the custom monitor
            CalculationOptions calcOptions = new CalculationOptions
            {
                CalculationMonitor = new CustomCalculationMonitor(),
                IgnoreError = false,
                Recursive = true
            };

            // Perform formula calculation; the monitor's BeforeCalculate will be invoked for each cell
            workbook.CalculateFormula(calcOptions);

            // Save the workbook (lifecycle rule: use provided saving logic)
            workbook.Save("CalculationMonitorDemo.xlsx");
        }
    }
}