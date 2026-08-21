// Title: C# – Create a Named Range for a Dynamic Array Spill and Use It in Subsequent Formulas with Aspose.Cells
// Description: Demonstrates how to set a dynamic array formula in Aspose.Cells, retrieve its spill area with GetArrayRange, define a workbook‑level named range that points to the spilled cells, and reference that name in later calculations such as SUM. The example builds the spill address, adds the name "MyArrayResult", recalculates formulas, and saves the workbook.
// Keywords: Aspose.Cells | C# | .NET | dynamic array formula | spill range | GetArrayRange | named range | workbook‑level name | SUM formula | Excel automation
// Common Searches: Aspose.Cells create named range from dynamic array spill | GetArrayRange example C# | how to reference spilled array result in Aspose.Cells | set dynamic array formula Aspose.Cells .NET | use named range in SUM formula Aspose.Cells
// Developer Intent: Define a named range that automatically references the result area of a dynamic array formula and reuse it in other calculations.
// Use Cases: Expose the spill area of a dynamic array as a reusable named range for aggregation functions. | Avoid hard‑coding cell addresses when the size of the array result may change. | Share the named range across multiple worksheets or external workbooks after saving.
// AI Prompts: Generate C# code that applies a dynamic array formula, obtains its spill range with GetArrayRange, and creates a global named range in Aspose.Cells. | Show how to use the named range created from a dynamic array spill in formulas like AVERAGE, MAX, or custom calculations. | Explain how to update the named range automatically when the source data changes and the spill area expands or contracts.

using System;
using Aspose.Cells;

namespace AsposeCellsArrayNamedRangeDemo
{
    // Demonstrates how to set a dynamic array formula in Aspose.Cells, retrieve its spill area with GetArrayRange, define a workbook‑level named range that points to the spilled cells, and reference that name in later calculations such as SUM. The example builds the spill address, adds the name "MyArrayResult", recalculates formulas, and saves the workbook.
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Populate some source data that will be used by the dynamic array formula
            //    Here we create a 3‑row, 2‑column matrix
            cells["A1"].PutValue(1);
            cells["A2"].PutValue(2);
            cells["A3"].PutValue(3);
            cells["B1"].PutValue(10);
            cells["B2"].PutValue(20);
            cells["B3"].PutValue(30);

            // 3. Set a dynamic array formula in cell C1.
            //    The formula multiplies the matrix A1:B3 by 2 and spills into neighboring cells.
            Cell formulaCell = cells["C1"];
            string arrayFormula = "=A1:B3*2";
            formulaCell.SetDynamicArrayFormula(arrayFormula, new FormulaParseOptions(), true);

            // 4. Refresh dynamic array formulas so that the spill range is materialized.
            //    The 'true' flag also calculates the values.
            workbook.RefreshDynamicArrayFormulas(true);

            // 5. Determine the actual spilled range of the dynamic array formula.
            //    GetArrayRange returns a CellArea describing the top‑left and bottom‑right cells.
            CellArea spillArea = formulaCell.GetArrayRange();

            // 6. Build the address string for the spilled range (e.g., "C1:D3").
            string startAddress = cells[spillArea.StartRow, spillArea.StartColumn].Name;
            string endAddress   = cells[spillArea.EndRow,   spillArea.EndColumn].Name;
            string spilledRangeAddress = $"{startAddress}:{endAddress}";

            // 7. Create a named range that refers to the spilled array result.
            //    The name will be scoped to the workbook (global name).
            int nameIndex = workbook.Worksheets.Names.Add("MyArrayResult");
            Name namedRange = workbook.Worksheets.Names[nameIndex];
            // RefersTo must start with '=' and include the sheet name.
            namedRange.RefersTo = $"={sheet.Name}!{spilledRangeAddress}";

            // 8. Use the named range in a subsequent calculation.
            //    For example, compute the sum of all values produced by the array formula.
            cells["E1"].Formula = "=SUM(MyArrayResult)";

            // 9. Calculate all formulas in the workbook.
            workbook.CalculateFormula();

            // 10. Output the result to the console (optional verification).
            Console.WriteLine($"Spilled range address: {spilledRangeAddress}");
            Console.WriteLine($"Sum of the array result (cell E1): {cells["E1"].Value}");

            // 11. Save the workbook.
            workbook.Save("ArrayNamedRangeDemo.xlsx");
        }
    }
}
