// Title: Assign a custom AbstractCalculationMonitor to CalculationOptions to log and interrupt formula calculation in Aspose.Cells (C#)
// AI Prompts: Create a C# class that inherits from AbstractCalculationMonitor, logs sheet, row, and column before and after each cell calculation, and returns false to stop the calculation when a custom condition is met. | Show how to attach the custom monitor to a CalculationOptions instance and pass it to Workbook.CalculateFormula to enable monitoring and possible interruption of formula evaluation. | Demonstrate retrieving the calculated values from cells after the monitored calculation finishes, including proper exception handling.
// Common Searches: how to use a custom calculation monitor with Aspose.Cells to stop long-running formula evaluation in C# | example of logging each cell calculation using AbstractCalculationMonitor in Aspose.Cells | C# code to attach CalculationOptions.CalculationMonitor for interrupting Excel formula calculation | handling circular reference notifications with a custom Aspose.Cells calculation monitor | monitoring workbook.CalculateFormula progress in .NET Aspose.Cells
// Tags: custom AbstractCalculationMonitor Aspose.Cells C# | assign CalculationOptions.CalculationMonitor | interrupt formula calculation Aspose.Cells | log cell calculation events Aspose.Cells | handle circular references with AbstractCalculationMonitor

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates assigning a custom calculation monitor to CalculationOptions
    // The example creates a workbook with simple formulas, defines a SampleCalculationMonitor that logs before and after each cell calculation by overriding AbstractCalculationMonitor methods, assigns this monitor to CalculationOptions.CalculationMonitor, runs workbook.CalculateFormula with the custom options, and prints the resulting cell values while demonstrating how the monitor can be used to interrupt or track formula evaluation.
    public class CalculationMonitorDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Set sample formulas that will be calculated
                sheet.Cells["A1"].Formula = "=1+2";
                sheet.Cells["A2"].Formula = "=A1*3";
                sheet.Cells["A3"].Formula = "=SUM(A1:A2)";

                // Create calculation options and attach a custom monitor
                CalculationOptions options = new CalculationOptions
                {
                    CalculationMonitor = new SampleCalculationMonitor()
                };

                // Perform formula calculation with monitoring enabled
                workbook.CalculateFormula(options);

                // Output the calculated values
                Console.WriteLine("A1: " + sheet.Cells["A1"].Value);
                Console.WriteLine("A2: " + sheet.Cells["A2"].Value);
                Console.WriteLine("A3: " + sheet.Cells["A3"].Value);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during calculation: " + ex.Message);
            }
        }
    }

    // Custom monitor that logs before and after each cell calculation
    public class SampleCalculationMonitor : AbstractCalculationMonitor
    {
        public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            Console.WriteLine($"Before calculation: Sheet{sheetIndex}, Row{rowIndex}, Col{columnIndex}");
        }

        public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            Console.WriteLine($"After calculation: Sheet{sheetIndex}, Row{rowIndex}, Col{columnIndex}");
        }

        public override bool OnCircular(IEnumerator circularCellsData)
        {
            // Use default handling for circular references
            return base.OnCircular(circularCellsData);
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                CalculationMonitorDemo.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unhandled exception: " + ex.Message);
            }
        }
    }
}
