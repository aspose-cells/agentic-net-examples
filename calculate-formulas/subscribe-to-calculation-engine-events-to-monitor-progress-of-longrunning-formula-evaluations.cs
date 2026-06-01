using System;
using System.Collections;
using Aspose.Cells;

public class CalculationMonitorDemo
{
    public static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some cells with values and formulas
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);
        sheet.Cells["A3"].Formula = "=A1+A2";          // Simple addition
        sheet.Cells["B1"].Formula = "=NOW()";          // Volatile function

        // Create calculation options and attach a custom monitor
        CalculationOptions options = new CalculationOptions
        {
            CalculationMonitor = new MyCalculationMonitor()
        };

        // Calculate all formulas while the monitor reports progress
        workbook.CalculateFormula(options);

        // Display calculated results
        Console.WriteLine("A3 = " + sheet.Cells["A3"].Value);
        Console.WriteLine("B1 = " + sheet.Cells["B1"].Value);

        // Save the workbook
        workbook.Save("CalculationMonitorDemo.xlsx");
    }

    // Custom monitor that logs before/after each cell calculation and handles circular references
    private class MyCalculationMonitor : AbstractCalculationMonitor
    {
        public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            Console.WriteLine($"Before calculation: Sheet {sheetIndex}, Row {rowIndex}, Column {columnIndex}");
        }

        public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            Console.WriteLine($"After calculation: Sheet {sheetIndex}, Row {rowIndex}, Column {columnIndex}");
            Console.WriteLine($"Original Value: {OriginalValue}, Calculated Value: {CalculatedValue}, Value Changed: {ValueChanged}");
        }

        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected:");
            while (circularCellsData.MoveNext())
            {
                Console.WriteLine($"  {circularCellsData.Current}");
            }
            // Continue calculation after reporting the circular reference
            return true;
        }
    }
}