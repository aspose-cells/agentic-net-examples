using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCircularReferenceDemo
{
    // Custom monitor to detect circular references during calculation
    public class CircularReferenceMonitor : AbstractCalculationMonitor
    {
        // Called when a circular reference is found
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected!");
            while (circularCellsData.MoveNext())
            {
                // Each item is a CalculationCell; its ToString gives the cell address
                Console.WriteLine($" - {circularCellsData.Current}");
            }
            // Stop further calculation of the circular cells
            return false;
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (no template file needed)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Set up a simple circular reference: A1 -> B1 -> A1
                sheet.Cells["A1"].Formula = "=B1";
                sheet.Cells["B1"].Formula = "=A1";

                // Prepare calculation options with the custom monitor
                CalculationOptions options = new CalculationOptions
                {
                    CalculationMonitor = new CircularReferenceMonitor(),
                    IgnoreError = true,
                    Recursive = true
                };

                // Perform calculation
                workbook.CalculateFormula(options);
                Console.WriteLine("Calculation completed without circular reference.");
                
                // Save the workbook
                string outputPath = "CircularReferenceDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (CellsException ex) // Aspose.Cells specific exception
            {
                Console.WriteLine($"CellsException caught: {ex.Message}");
            }
            catch (Exception ex) // Fallback for any other unexpected errors
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}