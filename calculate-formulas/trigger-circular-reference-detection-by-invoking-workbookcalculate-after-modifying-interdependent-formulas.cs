// Title: Detect Circular References in Aspose.Cells (C#) with a Custom AbstractCalculationMonitor
// Description: C# example that creates a workbook with interdependent formulas (A1 = B1, B1 = A1), implements a CircularReferenceMonitor by overriding AbstractCalculationMonitor.OnCircular, configures CalculationOptions (Recursive = true), runs Workbook.CalculateFormula, logs each cell in the loop, displays resulting values, and saves the file.
// Keywords: Aspose.Cells | circular reference detection | AbstractCalculationMonitor | Workbook.CalculateFormula | C# spreadsheet API | recursive calculation | formula monitoring | Excel circular reference handling
// Common Searches: Aspose.Cells detect circular reference C# | How to use AbstractCalculationMonitor in Aspose.Cells | Workbook.CalculateFormula with custom monitor | Enable recursive calculation Aspose.Cells | Log cells involved in circular reference Aspose
// Developer Intent: Identify and handle circular references during formula calculation using a custom monitor.
// Use Cases: Debug complex spreadsheets by listing every cell that participates in a circular reference. | Prevent infinite recursion by deciding whether to continue or abort calculation after detection. | Integrate automatic circular‑reference checks into a data‑processing pipeline before saving workbooks.
// AI Prompts: Write C# code that uses Aspose.Cells to detect circular references with a custom AbstractCalculationMonitor and optionally stop the calculation. | Show how to collect cell addresses from the OnCircular event and write them to a log file. | Explain the steps to configure CalculationOptions for recursive calculation and attach a custom monitor for circular reference detection.

using System;
using System.Collections;
using Aspose.Cells;

namespace CircularReferenceDemo
{
    // Custom monitor to handle circular reference events
    // C# example that creates a workbook with interdependent formulas (A1 = B1, B1 = A1), implements a CircularReferenceMonitor by overriding AbstractCalculationMonitor.OnCircular, configures CalculationOptions (Recursive = true), runs Workbook.CalculateFormula, logs each cell in the loop, displays resulting values, and saves the file.
    public class CircularReferenceMonitor : AbstractCalculationMonitor
    {
        // Called when a circular reference is detected during calculation
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected!");
            while (circularCellsData.MoveNext())
            {
                // Each item is a CalculationCell representing a cell involved in the circle
                Console.WriteLine($" - {circularCellsData.Current}");
            }
            // Return true to let the engine continue processing (or false to stop)
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
            Cells cells = sheet.Cells;

            // Set up interdependent formulas that create a circular reference
            cells["A1"].Formula = "=B1";
            cells["B1"].Formula = "=A1";

            // Prepare calculation options with the custom monitor
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = new CircularReferenceMonitor(),
                // Ensure recursive calculation is enabled (default true)
                Recursive = true
            };

            // Trigger calculation; this will invoke the monitor on circular detection
            workbook.CalculateFormula(options);

            // Output the values after calculation (they will remain unchanged or be marked as errors)
            Console.WriteLine($"A1 value after calculation: {cells["A1"].Value}");
            Console.WriteLine($"B1 value after calculation: {cells["B1"].Value}");

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("CircularReferenceDemo.xlsx");
        }
    }
}
