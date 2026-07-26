// Title: Add a Rectangle Shape Linked to a Cell and Capture Circular Reference Errors with a Custom Calculation Monitor (Aspose.Cells .NET)
// Description: Demonstrates how to create a workbook, insert a rectangle shape, link it to cell A1, set formulas that create a circular reference, and use a custom class derived from AbstractCalculationMonitor to log the involved cells. The example shows configuring CalculationOptions, handling the circular‑reference event, catching calculation exceptions, and saving the file.
// Keywords: Aspose.Cells add shape | link shape to cell | circular reference detection | custom calculation monitor | handle circular reference error | Aspose.Cells .NET example | shape linked cell formula | worksheet calculation monitor
// Common Searches: Aspose.Cells link shape to cell and detect circular reference | How to use AbstractCalculationMonitor for circular references | Add rectangle shape with linked cell in C# Aspose.Cells | Capture circular reference events during workbook calculation | Graceful error handling for circular formulas Aspose.Cells
// Developer Intent: Create a shape linked to a cell that participates in a circular reference and use a custom calculation monitor to detect and log the circular reference while handling any calculation errors.
// Use Cases: Display dynamic values on a shape that updates with cell changes and monitor for circular dependencies. | Implement a custom AbstractCalculationMonitor to enumerate and log cells involved in a circular reference. | Execute workbook calculations inside a try‑catch block to prevent crashes and still generate the output file.
// AI Prompts: Generate C# code that adds a circle shape linked to cell B2, creates a circular reference with another cell, and logs the reference using a custom calculation monitor in Aspose.Cells. | Show how to configure CalculationOptions with a custom monitor to capture circular reference events and continue processing after an exception. | Explain the parameters of SetLinkedCell for locale‑aware shape‑to‑cell linking in Aspose.Cells.

using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsCircularReferenceDemo
{
    // Custom calculation monitor to handle circular reference events
    // Demonstrates how to create a workbook, insert a rectangle shape, link it to cell A1, set formulas that create a circular reference, and use a custom class derived from AbstractCalculationMonitor to log the involved cells. The example shows configuring CalculationOptions, handling the circular‑reference event, catching calculation exceptions, and saving the file.
    class CircularReferenceMonitor : AbstractCalculationMonitor
    {
        // This method is called when the engine detects circular references
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected!");
            while (circularCellsData.MoveNext())
            {
                // Each item is a CalculationCell; display its address
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
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ---------- Add a rectangle shape ----------
            // Parameters: upper left row, upper left column, top, left, width, height
            Shape rect = sheet.Shapes.AddRectangle(2, 2, 0, 0, 120, 30);

            // Link the shape's value to cell A1
            rect.SetLinkedCell("$A$1", false, true); // absolute A1 reference, locale‑aware

            // ---------- Create a circular reference scenario ----------
            // A1 depends on B1 and B1 depends on A1
            sheet.Cells["A1"].Formula = "=B1";
            sheet.Cells["B1"].Formula = "=A1";

            // ---------- Set up calculation options with the custom monitor ----------
            CalculationOptions calcOptions = new CalculationOptions
            {
                CalculationMonitor = new CircularReferenceMonitor()
            };

            // ---------- Perform calculation and handle possible errors ----------
            try
            {
                workbook.CalculateFormula(calcOptions);
                Console.WriteLine("Calculation completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Calculation error: {ex.Message}");
            }

            // ---------- Save the workbook ----------
            workbook.Save("CircularReferenceDemo.xlsx");
        }
    }
}
