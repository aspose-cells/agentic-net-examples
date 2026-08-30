// Title: Subscribe to Aspose.Cells calculation engine events with a custom AbstractCalculationMonitor to log formula evaluation and handle circular references in C#
// AI Prompts: Create a class inheriting from AbstractCalculationMonitor, override BeforeCalculate, AfterCalculate, and OnCircular to write cell coordinates, original and new values, and return true for circular references, then assign it to CalculationOptions.CalculationMonitor and invoke Workbook.CalculateFormula. | Configure CalculationOptions with Recursive = true and a custom monitor to capture volatile function evaluation, run the calculation, and display the final values of the evaluated cells. | Demonstrate saving the workbook after the monitored calculation run, ensuring the custom monitor streams progress messages to the console throughout the process.
// Common Searches: Aspose.Cells C# monitor formula calculation progress with AbstractCalculationMonitor | How to detect circular references during Workbook.CalculateFormula in Aspose.Cells | Log before and after each cell calculation using Aspose.Cells CalculationOptions | Enable recursive calculation and volatile function tracking in Aspose.Cells .NET | Custom calculation monitor example for Aspose.Cells workbook evaluation
// Tags: custom AbstractCalculationMonitor implementation C# | Aspose.Cells calculation monitor logging | circular reference detection Aspose.Cells | recursive formula evaluation Aspose.Cells | volatile function handling Aspose.Cells

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsCalculationMonitorDemo
{
    // Custom monitor to track calculation progress
    // The example defines MyCalculationMonitor derived from AbstractCalculationMonitor to log cell coordinates, original and calculated values, and enumerate circular references. It attaches the monitor to CalculationOptions, enables recursive evaluation, runs Workbook.CalculateFormula, prints final results, and saves the workbook.
    public class MyCalculationMonitor : AbstractCalculationMonitor
    {
        // Called before each cell is calculated
        public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            Console.WriteLine($"Before: Sheet{sheetIndex}, Row{rowIndex}, Col{columnIndex}");
        }

        // Called after each cell is calculated
        public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            Console.WriteLine($"After : Sheet{sheetIndex}, Row{rowIndex}, Col{columnIndex}");
            Console.WriteLine($"  Original: {OriginalValue}, New: {CalculatedValue}, Changed: {ValueChanged}");
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
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data and formulas
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].Formula = "=A1+A2";          // Simple addition
            sheet.Cells["B1"].Formula = "=A3*2";           // Dependent on A3
            sheet.Cells["C1"].Formula = "=B1+NOW()";      // Volatile function

            // Set up calculation options with the custom monitor
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = new MyCalculationMonitor(),
                Recursive = true,
                IgnoreError = false
            };

            // Perform calculation; monitor callbacks will be invoked
            workbook.CalculateFormula(options);

            // Output final values
            Console.WriteLine("\nFinal Results:");
            Console.WriteLine($"A3 = {sheet.Cells["A3"].Value}");
            Console.WriteLine($"B1 = {sheet.Cells["B1"].Value}");
            Console.WriteLine($"C1 = {sheet.Cells["C1"].Value}");

            // Save the workbook (optional)
            workbook.Save("CalculationMonitorDemo.xlsx");
        }
    }
}
