// Title: How to extract the text of a complex array formula using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that assigns an array formula to a cell with Aspose.Cells and then uses the Formula property to obtain the exact formula string for debugging. | Show how to retrieve the formula text of a shared or non‑shared array formula in a worksheet cell using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# get formula string from cell containing array formula | How to read array formula text in a .NET workbook with Aspose.Cells | Debugging complex Excel array formulas by extracting formula text using Aspose.Cells API
// Tags: Aspose.Cells retrieve array formula text | C# extract cell formula string | Aspose.Cells debugging array formulas | SetArrayFormula and Formula property .NET | Excel array formula extraction with Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a workbook, applies the array formula SUM(IF(A1:A10>5, A1:A10, 0)) to cell B1, then reads the formula text via the cell's Formula property and writes it to the console.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Define a complex array formula (example: SUM of IF condition)
                string arrayFormula = "SUM(IF(A1:A10>5, A1:A10, 0))";

                // Get the target cell (B1)
                Cell cell = sheet.Cells["B1"];

                // Apply the array formula to the single cell.
                // Overload: SetArrayFormula(string formula, int firstRow, int firstColumn, bool isShared, bool isArray)
                cell.SetArrayFormula(arrayFormula, cell.Row, cell.Column, false, true);

                // Retrieve the formula text for debugging (Formula property works for array formulas)
                string formulaText = cell.Formula;

                // Output the formula text
                Console.WriteLine("Array formula text: " + formulaText);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}
