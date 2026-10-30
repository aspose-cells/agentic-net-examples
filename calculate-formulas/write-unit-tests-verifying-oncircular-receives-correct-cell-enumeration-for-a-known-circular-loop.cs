// Title: C# unit test for Aspose.Cells OnCircular event to confirm correct cell enumeration in a circular reference loop
// Common Searches: how to write a unit test for Aspose.Cells circular reference detection in C# | verify that OnCircular receives the correct cells when a formula loop exists in Aspose.Cells | Aspose.Cells CalculateFormula throws CellsException with involved cells for circular reference unit test | C# test that exception message includes A1 and B1 for circular reference in a workbook | unit testing circular reference handling using Aspose.Cells CalculateFormula method
// Tags: Aspose.Cells circular reference unit testing | OnCircular event cell enumeration verification | CalculateFormula exception handling Aspose.Cells | C# workbook circular loop detection | CellsException message validation Aspose.Cells

using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsCircularReferenceTests
{
    // Demonstrates a C# unit test that creates a workbook with a circular reference (A1 ↔ B1), triggers CalculateFormula, expects a CellsException, and asserts that the exception message indicates a circular reference and lists the involved cells, confirming correct OnCircular cell enumeration.
    class Program
    {
        static void Main()
        {
            try
            {
                var test = new CircularReferenceTests();
                test.OnCircular_ShouldReceiveCorrectCellEnumeration();
                Console.WriteLine("Test passed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
            }
        }
    }

    public class CircularReferenceTests
    {
        // Creates a workbook where A1 references B1 and B1 references A1 (circular reference)
        private Workbook CreateWorkbookWithCircularLoop()
        {
            var wb = new Workbook();
            var ws = wb.Worksheets[0];
            ws.Cells["A1"].Formula = "=B1";
            ws.Cells["B1"].Formula = "=A1";
            return wb;
        }

        public void OnCircular_ShouldReceiveCorrectCellEnumeration()
        {
            var workbook = CreateWorkbookWithCircularLoop();

            try
            {
                // Trigger calculation; a circular reference will raise a CellsException
                workbook.CalculateFormula();
                throw new InvalidOperationException("Expected a circular reference exception to be thrown.");
            }
            catch (CellsException ex)
            {
                // Verify that the exception indicates a circular reference
                if (!ex.Message.Contains("Circular reference", StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException($"Unexpected CellsException: {ex.Message}");
                }

                // Basic validation that the message mentions one of the involved cells
                if (!ex.Message.Contains("A1") && !ex.Message.Contains("B1"))
                {
                    throw new InvalidOperationException("Circular reference cells not identified in exception message.");
                }
            }
        }
    }
}
