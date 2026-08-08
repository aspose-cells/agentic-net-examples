// Title: Aspose.Cells C# Example: IFNA Formula with Workbook.CalculateFormula
// Description: Demonstrates how to set an IFNA formula in cell B1, clear or populate cell A1, recalculate the workbook with Workbook.CalculateFormula, and read the fallback string or numeric result using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | IFNA function | C# | .NET | Workbook.CalculateFormula | formula evaluation | fallback value | Excel automation | clear cell content | unit test example
// Common Searches: Aspose.Cells IFNA example C# | how to use Workbook.CalculateFormula | IFNA fallback string Aspose.Cells | recalculate workbook after changing cell value | read formula result Aspose.Cells C#
// Developer Intent: Show how IFNA returns a default value when the referenced cell is empty and returns the cell's actual value when it contains data, using Aspose.Cells formula calculation.
// Use Cases: Validate IFNA handling of missing or numeric data in generated Excel reports. | Replace empty cells with a default string while programmatically building workbooks. | Create automated tests that modify a source cell, invoke CalculateFormula, and verify the IFNA result.
// AI Prompts: Generate a C# snippet that inserts an IFNA formula with Aspose.Cells, runs Workbook.CalculateFormula, and prints the result. | Explain how to clear a cell's contents in Aspose.Cells before recalculating an IFNA formula. | Provide a unit‑test code sample that asserts IFNA returns the fallback string and then the actual value after updating the source cell.

using System;
using Aspose.Cells;

namespace AsposeCellsIFNADemo
{
    // Demonstrates how to set an IFNA formula in cell B1, clear or populate cell A1, recalculate the workbook with Workbook.CalculateFormula, and read the fallback string or numeric result using Aspose.Cells for .NET.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Scenario 1: A1 is empty, IFNA should return the fallback string
                // Clear A1 by putting an empty string (Aspose.Cells has no ClearContents method)
                cells["A1"].PutValue(string.Empty);
                cells["B1"].Formula = "=IFNA(A1, \"fallback\")";

                // Calculate all formulas in the workbook
                workbook.CalculateFormula();

                // Display the result of IFNA when A1 is empty
                Console.WriteLine("B1 (A1 empty) = " + cells["B1"].StringValue); // Expected: fallback

                // Scenario 2: A1 contains a numeric value, IFNA should return that value
                cells["A1"].PutValue(42);

                // Recalculate after changing the input value
                workbook.CalculateFormula();

                // Display the result of IFNA when A1 has a value
                Console.WriteLine("B1 (A1 = 42) = " + cells["B1"].StringValue); // Expected: 42

                // (Optional) Save the workbook to verify the formula and results in Excel
                // workbook.Save("IFNADemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
