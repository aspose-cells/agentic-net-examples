// Title: C# – Insert a SORT Dynamic Array Formula in B2 and Verify the Spill Range with Aspose.Cells
// Description: Creates a workbook, fills A1:A5 with unsorted numbers, applies the SORT dynamic‑array formula (=SORT(A1:A5,1,-1)) to B2 using SetDynamicArrayFormula, recalculates, refreshes dynamic arrays, retrieves both the expected and actual spill areas via GetArrayRange, prints the sorted values, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | SORT formula | dynamic array | SetDynamicArrayFormula | GetArrayRange | RefreshDynamicArrayFormulas | spill range | Excel sorting | B2 cell | A1:A5 data
// Common Searches: Aspose.Cells set SORT dynamic array formula C# | how to get spill range of a dynamic array formula Aspose.Cells | RefreshDynamicArrayFormulas after calculation .NET | verify SORT formula spill area in Excel using Aspose | C# example for SetDynamicArrayFormula and GetArrayRange
// Developer Intent: Add a SORT dynamic‑array formula to cell B2, then confirm that the formula’s spilled range matches the expected area returned by SetDynamicArrayFormula.
// Use Cases: Sort a list of values in a worksheet without manual formulas and capture the exact spill region. | Programmatically recalculate a workbook and ensure dynamic‑array results are up‑to‑date before reading them. | Iterate over the spilled cells returned by GetArrayRange to process or export sorted data.
// AI Prompts: Write C# code using Aspose.Cells to insert a SORT dynamic‑array formula in B2, refresh formulas, and compare expected vs. actual spill ranges. | Explain the role of GetArrayRange after inserting a SORT formula and how obstacles affect the spill area. | Provide a step‑by‑step guide to validate that a SORT dynamic‑array formula spills correctly by enumerating the spilled cells and outputting their values.

using System;
using Aspose.Cells;

// Creates a workbook, fills A1:A5 with unsorted numbers, applies the SORT dynamic‑array formula (=SORT(A1:A5,1,-1)) to B2 using SetDynamicArrayFormula, recalculates, refreshes dynamic arrays, retrieves both the expected and actual spill areas via GetArrayRange, prints the sorted values, and saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        Cells cells = ws.Cells;

        // Populate unsorted data in column A (A1:A5)
        cells["A1"].PutValue(5);
        cells["A2"].PutValue(2);
        cells["A3"].PutValue(8);
        cells["A4"].PutValue(1);
        cells["A5"].PutValue(4);

        // Set a SORT dynamic array formula in cell B2
        // The formula sorts the range A1:A5 in descending order
        Cell targetCell = cells["B2"];
        string sortFormula = "=SORT(A1:A5,1,-1)";

        // SetDynamicArrayFormula returns the range that the formula should spill into
        CellArea expectedSpill = targetCell.SetDynamicArrayFormula(sortFormula, new FormulaParseOptions(), true);

        // Calculate formulas and refresh dynamic array formulas to ensure the spill range is up‑to‑date
        wb.CalculateFormula();
        wb.RefreshDynamicArrayFormulas(true);

        // Get the actual spilled range using GetArrayRange (may differ if there are obstacles)
        CellArea actualSpill = targetCell.GetArrayRange();

        // Output the expected and actual spill ranges
        Console.WriteLine($"Expected spill range (from SetDynamicArrayFormula): Rows {expectedSpill.StartRow}-{expectedSpill.EndRow}, Columns {expectedSpill.StartColumn}-{expectedSpill.EndColumn}");
        Console.WriteLine($"Actual spill range (from GetArrayRange): Rows {actualSpill.StartRow}-{actualSpill.EndRow}, Columns {actualSpill.StartColumn}-{actualSpill.EndColumn}");

        // Verify the spilled values by iterating over the actual spill range
        Console.WriteLine("Spilled values:");
        for (int r = actualSpill.StartRow; r <= actualSpill.EndRow; r++)
        {
            for (int c = actualSpill.StartColumn; c <= actualSpill.EndColumn; c++)
            {
                Console.Write(cells[r, c].Value + "\t");
            }
            Console.WriteLine();
        }

        // Save the workbook to verify the result in Excel
        wb.Save("SortDynamicArrayDemo.xlsx");
    }
}
