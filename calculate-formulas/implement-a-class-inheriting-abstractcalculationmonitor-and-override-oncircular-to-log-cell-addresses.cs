// Title: Implement a custom AbstractCalculationMonitor in Aspose.Cells (.NET) to log circular reference cell addresses during workbook formula calculation
// AI Prompts: Write a C# class that inherits from Aspose.Cells.AbstractCalculationMonitor and overrides the OnCircular method to print each circular reference cell's address to the console. | Show how to assign the custom monitor to CalculationOptions and invoke Workbook.CalculateFormula so that circular references are detected and logged.
// Common Searches: how to log circular reference cells using Aspose.Cells calculation monitor in C# | Aspose.Cells detect circular references during formula calculation .NET example | override OnCircular method to output cell addresses in Aspose.Cells workbook
// Tags: Aspose.Cells calculation monitor implementation | Aspose.Cells circular reference handling | circular reference callback implementation | formula calculation with monitor

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsCircularReferenceDemo
{
    // Custom monitor that logs circular reference cell addresses
    // The example defines a CircularReferenceMonitor class that inherits from AbstractCalculationMonitor and overrides OnCircular to write each involved CalculationCell to the console. The monitor is attached via CalculationOptions, a workbook with a simple A1↔B1 circular reference is created, Workbook.CalculateFormula triggers the monitor, and the workbook is saved.
    public class CircularReferenceMonitor : AbstractCalculationMonitor
    {
        // This method is called when the calculation engine detects circular references.
        // The IEnumerator contains CalculationCell objects representing the cells involved.
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected. Involved cells:");

            // Iterate through the circular cells and output their addresses.
            while (circularCellsData.MoveNext())
            {
                // The items are CalculationCell objects; their ToString() provides a readable description.
                Console.WriteLine($"- {circularCellsData.Current}");
            }

            // Return true to let the engine continue processing (or false to stop).
            return true;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set up a simple circular reference: A1 -> B1 -> A1
            sheet.Cells["A1"].Formula = "=B1";
            sheet.Cells["B1"].Formula = "=A1";

            // Configure calculation options to use our custom monitor.
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = new CircularReferenceMonitor()
            };

            // Perform formula calculation. The monitor will be invoked for the circular reference.
            workbook.CalculateFormula(options);

            // Save the workbook (optional, just to demonstrate full lifecycle).
            workbook.Save("CircularReferenceDemo.xlsx");
        }
    }
}
