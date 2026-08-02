// Title: Detect Circular References in Aspose.Cells (.NET) Using a Custom CalculationMonitor
// Description: Learn how to identify and handle circular references in an Aspose.Cells workbook by implementing a custom CalculationMonitor (derived from AbstractCalculationMonitor) and invoking Workbook.CalculateFormula with CalculationOptions. The example logs each cell involved in the loop and shows how to continue or abort the calculation.
// Keywords: Aspose.Cells circular reference | Custom CalculationMonitor .NET | AbstractCalculationMonitor example | Workbook.CalculateFormula circular detection | C# Excel formula engine | Aspose.Cells error handling | Excel circular reference detection | Aspose.Cells calculation options | Formula evaluation Aspose.Cells | C# Aspose.Cells monitor
// Common Searches: Aspose.Cells detect circular reference C# | How to use CalculationMonitor for circular formulas | Workbook.CalculateFormula with custom monitor | Capture circular reference events Aspose.Cells | Implement AbstractCalculationMonitor in .NET
// Developer Intent: The developer needs to detect circular references during formula calculation and process the event with custom logic.
// Use Cases: Log every cell that participates in a circular reference before saving the workbook. | Abort the calculation engine when a circular reference is found and replace the cells with a predefined error value. | Show user-friendly messages by extracting circular reference details via a custom CalculationMonitor.
// AI Prompts: Generate C# code that uses Aspose.Cells to detect circular references with a custom AbstractCalculationMonitor and logs each cell address. | Modify the OnCircular method to stop the calculation and assign a custom error value to the cells involved in the loop. | Explain how to configure CalculationOptions to attach a custom monitor and retrieve circular reference information after calling Workbook.CalculateFormula.

using System;
using System.Collections;
using Aspose.Cells;

namespace CircularReferenceDemo
{
    // Custom monitor to handle circular reference events
    // Learn how to identify and handle circular references in an Aspose.Cells workbook by implementing a custom CalculationMonitor (derived from AbstractCalculationMonitor) and invoking Workbook.CalculateFormula with CalculationOptions. The example logs each cell involved in the loop and shows how to continue or abort the calculation.
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
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set up interdependent formulas that create a circular reference
            cells["A1"].Formula = "=B1";
            cells["B1"].Formula = "=A1";

            // Prepare calculation options with the custom monitor
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = new CircularReferenceMonitor()
            };

            // Trigger calculation; this will invoke the monitor on circular detection
            workbook.CalculateFormula(options);

            // Output the resulting values (they will be #REF! or similar)
            Console.WriteLine($"A1 value after calculation: {cells["A1"].Value}");
            Console.WriteLine($"B1 value after calculation: {cells["B1"].Value}");

            // Save the workbook to verify the formulas are stored
            workbook.Save("CircularReferenceDemo.xlsx");
        }
    }
}
