// Title: Assign a Custom Calculation Monitor to Workbook.Settings.CalculationMonitor in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a class derived from AbstractCalculationMonitor, attach it to Workbook.Settings.CalculationMonitor, and run Workbook.CalculateFormula. The monitor logs cell coordinates before and after each calculation, reports original and new values, and stops processing when a circular reference is detected.
// Keywords: Aspose.Cells | C# | .NET | custom calculation monitor | AbstractCalculationMonitor | Workbook.Settings.CalculationMonitor | formula calculation callbacks | circular reference detection | Excel automation logging | debugging formula evaluation
// Common Searches: how to use a custom calculation monitor in Aspose.Cells C# | set Workbook.Settings.CalculationMonitor before loading workbook | log formula calculation steps with Aspose.Cells | stop calculation on circular reference Aspose.Cells | Aspose.Cells calculation monitor example
// Developer Intent: Implement a custom monitor to receive before/after calculation events and to intercept circular references during Excel formula evaluation with Aspose.Cells.
// Use Cases: Debug complex workbooks by tracing each cell's calculation order. | Create an audit trail that records original and newly calculated values. | Prevent infinite loops by aborting calculation when a circular reference is found.
// AI Prompts: Show C# code that assigns MyCalculationMonitor to Workbook.Settings.CalculationMonitor before loading an Excel file with Aspose.Cells. | Explain how to log cell coordinates and value changes using AbstractCalculationMonitor in Aspose.Cells. | Compare using CalculationOptions.CalculationMonitor versus Workbook.Settings.CalculationMonitor for formula evaluation.

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCustomMonitorDemo
{
    // Custom monitor implementation
    // Demonstrates how to create a class derived from AbstractCalculationMonitor, attach it to Workbook.Settings.CalculationMonitor, and run Workbook.CalculateFormula. The monitor logs cell coordinates before and after each calculation, reports original and new values, and stops processing when a circular reference is detected.
    public class MyCalculationMonitor : AbstractCalculationMonitor
    {
        public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            Console.WriteLine($"Before calculate - Sheet:{sheetIndex} Row:{rowIndex} Col:{columnIndex}");
        }

        public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            Console.WriteLine($"After calculate  - Sheet:{sheetIndex} Row:{rowIndex} Col:{columnIndex}");
            Console.WriteLine($"  Original: {OriginalValue}, New: {CalculatedValue}, Changed: {ValueChanged}");
        }

        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected.");
            // Return false to stop calculation when a circular reference is found
            return false;
        }
    }

    class Program
    {
        static void Main()
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Create the custom monitor instance
                MyCalculationMonitor monitor = new MyCalculationMonitor();

                // Prepare calculation options and assign the monitor
                CalculationOptions calcOptions = new CalculationOptions
                {
                    CalculationMonitor = monitor
                };

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Perform formula calculation using the options that contain the monitor
                workbook.CalculateFormula(calcOptions);

                // Save the workbook after calculation
                workbook.Save(outputPath);
                Console.WriteLine($"Calculation completed. Output saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
