// Title: Create a Named Range from a Dynamic Array Spill and Use It in Calculations – Aspose.Cells C#
// Description: Demonstrates how to set a dynamic array formula in a workbook, retrieve its spill area with GetArrayRange, define a named range that points to the spilled cells, and reference that name in subsequent formulas such as SUM. The example populates sample data, refreshes the array, outputs the spill values, calculates the total, and saves the file.
// Keywords: Aspose.Cells | C# | dynamic array formula | spill range | GetArrayRange | named range | SetDynamicArrayFormula | SUM formula | workbook automation | Excel array spill
// Common Searches: Aspose.Cells create named range from array spill | GetArrayRange dynamic array C# | use spilled array as named range Aspose.Cells | sum named range that references dynamic array | refresh dynamic array formulas Aspose.Cells
// Developer Intent: Define a named range that references the result of a dynamic array formula and use that range in later calculations.
// Use Cases: Generate a dynamic array (e.g., =A1:A3*2) and capture its spill area programmatically. | Create a workbook‑level named range that points to the spilled cells for reuse. | Apply the named range in other formulas such as SUM, AVERAGE, or custom calculations. | Validate the spill values and aggregated results before saving the workbook.
// AI Prompts: Write C# code with Aspose.Cells to set a dynamic array formula, obtain its spill range, create a named range for that range, and calculate the sum of the named range. | Explain how GetArrayRange and Workbook.Worksheets.Names work together to build a named range from a spilled array in Aspose.Cells. | Provide a step‑by‑step example that creates a named range from a dynamic array result, uses it in another formula, and saves the workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsArrayNamedRangeDemo
{
    // Demonstrates how to set a dynamic array formula in a workbook, retrieve its spill area with GetArrayRange, define a named range that points to the spilled cells, and reference that name in subsequent formulas such as SUM. The example populates sample data, refreshes the array, outputs the spill values, calculates the total, and saves the file.
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Populate some sample data that will be used by the array formula
            //    A1:A3 will contain numbers 1, 2, 3
            cells["A1"].PutValue(1);
            cells["A2"].PutValue(2);
            cells["A3"].PutValue(3);

            // 3. Set a dynamic array formula in cell B1.
            //    The formula multiplies each value in A1:A3 by 2 and spills the results.
            Cell formulaCell = cells["B1"];
            formulaCell.SetDynamicArrayFormula("=A1:A3*2", new FormulaParseOptions(), true);

            // 4. Refresh dynamic array formulas so that the spill range is materialized
            workbook.RefreshDynamicArrayFormulas(true);

            // 5. Determine the spilled range of the dynamic array formula
            //    GetArrayRange returns the area that the formula occupies.
            CellArea spillArea = formulaCell.GetArrayRange();

            // 6. Build the address string for the spilled range (e.g., Sheet1!$B$1:$B$3)
            string startCell = cells[spillArea.StartRow, spillArea.StartColumn].Name;
            string endCell   = cells[spillArea.EndRow,   spillArea.EndColumn].Name;
            string refersTo  = $"={sheet.Name}!{startCell}:{endCell}";

            // 7. Create a named range that refers to the spilled array result
            int nameIndex = workbook.Worksheets.Names.Add("MyArray");
            Name namedRange = workbook.Worksheets.Names[nameIndex];
            namedRange.RefersTo = refersTo;   // e.g., =Sheet1!$B$1:$B$3

            // 8. Use the named range in a subsequent calculation (sum of the array)
            cells["C1"].Formula = "=SUM(MyArray)";

            // 9. Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // 10. Output the results to the console for verification
            Console.WriteLine("Spill range address: " + refersTo);
            Console.WriteLine("Values in spilled range:");
            for (int r = spillArea.StartRow; r <= spillArea.EndRow; r++)
            {
                for (int c = spillArea.StartColumn; c <= spillArea.EndColumn; c++)
                {
                    Console.Write(cells[r, c].Value + "\t");
                }
                Console.WriteLine();
            }
            Console.WriteLine("Sum of named range (MyArray) in C1: " + cells["C1"].Value);

            // 11. Save the workbook
            workbook.Save("ArrayNamedRangeDemo.xlsx");
        }
    }
}
