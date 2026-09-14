// Title: How to Detect and Log Circular References in Excel Formulas Using Aspose.Cells C# Calculation Monitor
// AI Prompts: Create a C# class that inherits from AbstractCalculationMonitor and writes each cell address involved in a circular reference to the console. | Configure CalculationOptions with the custom monitor and enable recursive calculation to capture circular formula loops during workbook.CalculateFormula. | Adjust the monitor to abort the calculation by returning false when a circular reference is encountered.
// Common Searches: Aspose.Cells C# example for detecting circular references during formula calculation | log cells that cause circular reference error using Aspose.Cells calculation monitor | how to use AbstractCalculationMonitor to handle circular formulas in a workbook | C# workbook.CalculateFormula circular reference detection Aspose.Cells
// Tags: Aspose.Cells custom calculation monitor | circular reference detection Aspose.Cells | log offending cells C# Aspose.Cells | recursive formula calculation Aspose.Cells | Excel workbook circular formula handling

using System;
using System.Collections;
using Aspose.Cells;

namespace CircularReferenceDetectionDemo
{
    // Custom monitor to capture circular reference information
    // Demonstrates a custom CircularReferenceMonitor derived from AbstractCalculationMonitor that logs the addresses of cells participating in a circular reference (e.g., A1↔B1) during workbook.CalculateFormula, with CalculationOptions configured for recursive evaluation and optional continuation or abort behavior.
    public class CircularReferenceMonitor : AbstractCalculationMonitor
    {
        // This method is called when the calculation engine detects a circular reference
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected! Offending cells:");

            // Iterate through the cells involved in the circular reference
            while (circularCellsData.MoveNext())
            {
                // The enumerated object represents a cell participating in the circular loop.
                // Its ToString() implementation provides a readable address (e.g., "Sheet1!A1").
                Console.WriteLine($" - {circularCellsData.Current}");
            }

            // Return true to let the engine continue processing (or false to stop).
            // Here we allow the engine to continue with default handling.
            return true;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set up a simple circular reference scenario:
            // A1 depends on B1, and B1 depends on A1
            cells["A1"].Formula = "=B1";
            cells["B1"].Formula = "=A1";

            // Configure calculation options with the custom monitor
            CalculationOptions calcOptions = new CalculationOptions
            {
                CalculationMonitor = new CircularReferenceMonitor(),
                // Optional: you can control recursion or iterative calculation here
                Recursive = true
            };

            // Perform formula calculation; the monitor will be invoked automatically
            workbook.CalculateFormula(calcOptions);

            // Save the workbook (optional, just to demonstrate lifecycle compliance)
            workbook.Save("CircularReferenceDemo.xlsx");
        }
    }
}
