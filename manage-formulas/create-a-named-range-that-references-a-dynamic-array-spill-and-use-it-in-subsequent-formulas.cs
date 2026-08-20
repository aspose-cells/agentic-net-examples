// Title: Create a named range for a dynamic array spill and use it in formulas – Aspose.Cells for .NET
// Description: This C# example shows how to set a dynamic array formula with SetDynamicArrayFormula, capture the resulting spill CellArea, convert it to a Range, add a worksheet‑level named range that points to the spill, and reference that name in another formula (e.g., SUM). The workbook is refreshed, calculated, and saved.
// Keywords: Aspose.Cells, .NET, C#, dynamic array, spill range, named range, SetDynamicArrayFormula, SEQUENCE function, workbook refresh, CalculateFormula, Excel formulas, worksheet names
// Common Searches: Aspose.Cells create named range from dynamic array spill | SetDynamicArrayFormula example C# | reference spilled array in another formula Aspose.Cells | how to refresh dynamic array formulas Aspose.Cells | C# Excel dynamic array named range
// Developer Intent: Define a named range that automatically tracks a dynamic array spill and reuse it in subsequent calculations.
// Use Cases: Generate a sequence whose length is driven by a cell value and sum the results via a named range. | Keep a named range synchronized with a changing spill size after updating the source cell. | Export a workbook that contains a dynamic array and a reusable named range for downstream reporting.
// AI Prompts: Show how to update the named range automatically when the value in B1 changes. | Give an example of using the named range MySpill in a VLOOKUP or INDEX/MATCH formula. | Explain how to obtain the address of the spilled range without creating a Range object.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsDynamicArrayNamedRangeDemo
{
    // This C# example shows how to set a dynamic array formula with SetDynamicArrayFormula, capture the resulting spill CellArea, convert it to a Range, add a worksheet‑level named range that points to the spill, and reference that name in another formula (e.g., SUM). The workbook is refreshed, calculated, and saved.
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // 2. Put a value that will control the size of the dynamic array
                //    The SEQUENCE function will generate numbers from 1 to the value in B1
                cells["B1"].PutValue(5);

                // 3. Set a dynamic array formula in A1.
                //    The method returns the CellArea that the formula is expected to spill into.
                Cell spillCell = cells["A1"];
                CellArea spillArea = spillCell.SetDynamicArrayFormula(
                    "=SEQUENCE(B1)",          // array formula
                    new FormulaParseOptions(), // parsing options (default)
                    true);                    // calculate the values immediately

                // 4. Refresh dynamic array formulas to ensure the spill range is up‑to‑date
                //    (necessary if the source data changes after the formula is set)
                workbook.RefreshDynamicArrayFormulas(true);

                // 5. Convert the returned CellArea to a Range object
                int rows = spillArea.EndRow - spillArea.StartRow + 1;
                int cols = spillArea.EndColumn - spillArea.StartColumn + 1;
                AsposeRange spillRange = cells.CreateRange(
                    spillArea.StartRow,
                    spillArea.StartColumn,
                    rows,
                    cols);

                // 6. Create a named range that refers to the spilled dynamic array
                int nameIndex = workbook.Worksheets.Names.Add("MySpill");
                Name namedRange = workbook.Worksheets.Names[nameIndex];
                // The RefersTo string must start with '=' and contain the full address of the range
                namedRange.RefersTo = "=" + spillRange.RefersTo;

                // 7. Use the named range in another formula (e.g., sum the spilled values)
                cells["C1"].Formula = "=SUM(MySpill)";

                // 8. Calculate all formulas in the workbook
                workbook.CalculateFormula();

                // 9. Output the results to the console (optional verification)
                Console.WriteLine("Dynamic array spill values:");
                for (int r = spillArea.StartRow; r <= spillArea.EndRow; r++)
                {
                    Console.WriteLine(cells[r, spillArea.StartColumn].Value);
                }
                Console.WriteLine($"Sum of spilled values (MySpill) = {cells["C1"].Value}");

                // 10. Save the workbook
                workbook.Save("DynamicArrayNamedRangeDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
