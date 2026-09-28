// Title: Detect circular formulas in Aspose.Cells using Workbook.CalculateFormula with a custom CalculationMonitor (C#)
// AI Prompts: Write C# code that creates a workbook, sets cell A1 to =B1 and B1 to =A1, defines a class inheriting from AbstractCalculationMonitor that records circular cells, configures CalculationOptions with this monitor, and calls Workbook.CalculateFormula to capture the circular condition. | Show how to override the OnCircular method in a custom monitor to iterate over the circularCellsData collection and output each cell address while Aspose.Cells evaluates interdependent formulas.
// Common Searches: example of using Aspose.Cells CalculationMonitor to log circular formula errors in C# | trigger formula calculation after creating interdependent cells with Aspose.Cells | C# Aspose.Cells tutorial for handling circular reference events | steps to configure CalculationOptions for circular formula monitoring in Aspose.Cells
// Tags: Aspose.Cells calculateformula circular formula monitor | C# custom calculation monitor Aspose.Cells | interdependent formulas detection Aspose.Cells | Workbook.CalculateFormula with custom monitor | circular formula handling Aspose.Cells

using System;
using System.Collections;
using Aspose.Cells;

namespace CircularReferenceDemo
{
    // The sample creates a new workbook, assigns A1 =B1 and B1 =A1 to form a circular reference, implements a CircularReferenceMonitor derived from AbstractCalculationMonitor that logs each cell involved, attaches the monitor via CalculationOptions, invokes Workbook.CalculateFormula to trigger detection, prints the circular cells to the console, and saves the file as CircularReferenceDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set up interdependent formulas to create a circular reference
            sheet.Cells["A1"].Formula = "=B1";
            sheet.Cells["B1"].Formula = "=A1";

            // Configure calculation options with a custom monitor
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = new CircularReferenceMonitor()
            };

            // Trigger calculation; the monitor will detect the circular reference
            workbook.CalculateFormula(options);

            // Save the workbook (optional)
            workbook.Save("CircularReferenceDemo.xlsx");
        }

        // Custom monitor that handles circular reference events
        private class CircularReferenceMonitor : AbstractCalculationMonitor
        {
            public override bool OnCircular(IEnumerator circularCellsData)
            {
                Console.WriteLine("Circular reference detected:");
                while (circularCellsData.MoveNext())
                {
                    // Each item is a CalculationCell; display its address or description
                    Console.WriteLine(circularCellsData.Current);
                }
                // Return true to let the engine continue processing the circular cells
                return true;
            }
        }
    }
}
