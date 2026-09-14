// Title: Implement a custom AbstractCalculationMonitor in C# to log cell coordinates and detect circular references during Aspose.Cells formula calculation
// AI Prompts: Create a C# class that inherits from AbstractCalculationMonitor and overrides BeforeCalculate to write the sheet index, row index, and column index of each cell to the console. | Add overrides for AfterCalculate and OnCircular in the same class to output value changes and report circular reference detection. | Configure CalculationOptions with the custom monitor, invoke Workbook.CalculateFormula, and save the workbook to demonstrate the monitoring workflow.
// Common Searches: how to log each cell address before formula evaluation with Aspose.Cells C# | using AbstractCalculationMonitor to monitor formula calculation in Aspose.Cells | detect and handle circular references during Aspose.Cells calculation with a custom monitor | configure CalculationOptions to attach a custom calculation monitor in Aspose.Cells | example of overriding BeforeCalculate in Aspose.Cells for debugging formulas
// Tags: Aspose.Cells calculation monitor implementation | cell address logging during formula evaluation C# | value change detection after calculation Aspose.Cells | circular reference detection with OnCircular Aspose.Cells | configure CalculationOptions with custom monitor Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsCalculationMonitorDemo
{
    // Custom monitor that inspects each cell before it is calculated
    // The example defines MyCalculationMonitor, a class derived from AbstractCalculationMonitor that overrides BeforeCalculate to output sheet, row, and column indices before each cell is calculated, optionally logs value changes in AfterCalculate and reports circular references in OnCircular. The program creates a workbook, adds values and formulas, sets up CalculationOptions with this monitor, runs Workbook.CalculateFormula, and saves the workbook, demonstrating how to monitor and debug formula calculations in Aspose.Cells.
    public class MyCalculationMonitor : AbstractCalculationMonitor
    {
        // Called before a cell is calculated
        public override void BeforeCalculate(int sheetIndex, int rowIndex, int colIndex)
        {
            // Retrieve the workbook, worksheet and cell using the indices
            // Note: The monitor does not have direct access to the workbook,
            // so we rely on the static Workbook object created in the demo.
            // For demonstration we simply output the cell location and its original value.
            Console.WriteLine($"Before calculating Sheet:{sheetIndex} Row:{rowIndex} Column:{colIndex}");
        }

        // Optional: after calculation we can see if the value changed
        public override void AfterCalculate(int sheetIndex, int rowIndex, int colIndex)
        {
            if (ValueChanged)
            {
                Console.WriteLine($"Cell changed. Original: {OriginalValue}, New: {CalculatedValue}");
            }
        }

        // Optional: handle circular references
        public override bool OnCircular(System.Collections.IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected.");
            return false; // stop further calculation for circular cells
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data and formulas
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].Formula = "=A1+A2";   // Simple addition
            sheet.Cells["B1"].Formula = "=A3*2";    // Dependent formula

            // Instantiate the custom calculation monitor
            MyCalculationMonitor monitor = new MyCalculationMonitor();

            // Configure calculation options to use the monitor
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = monitor,
                IgnoreError = false,
                Recursive = true
            };

            // Perform formula calculation with monitoring
            workbook.CalculateFormula(options);

            // Save the workbook (demonstrates usage of the save lifecycle)
            workbook.Save("MonitoredCalculation.xlsx");
        }
    }
}
