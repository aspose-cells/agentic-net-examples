// Title: Aspose.Cells for .NET: Create a Named Range from a Dynamic Array Spill and Use It in Formulas (C#)
// Description: Learn how to set a SEQUENCE dynamic‑array formula, refresh the spill, capture its CellArea, define a workbook named range that points to the spilled cells, and reference that name in a SUM (or other) formula using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | dynamic array | spill range | named range | SEQUENCE function | RefreshDynamicArrayFormulas | CellArea | SUM formula | Excel automation | workbook naming
// Common Searches: Aspose.Cells create named range from dynamic array spill | C# set SEQUENCE formula and get spill area | RefreshDynamicArrayFormulas Aspose.Cells example | How to reference spilled array with a name in Aspose.Cells | Sum dynamic array spill using named range C#
// Developer Intent: Generate a named range that automatically tracks the spill area of a dynamic‑array formula and reuse that name in subsequent calculations.
// Use Cases: Create a SEQUENCE spill, capture its CellArea, and expose it as a named range for downstream formulas. | Replace hard‑coded cell references with a dynamic named range to calculate SUM, AVERAGE, COUNT, etc., on a spill that can change size. | Persist a workbook containing a dynamic array and a named range so other tools or users can reference the spill by name.
// AI Prompts: Show C# code with Aspose.Cells that sets a SEQUENCE dynamic array, refreshes the spill, creates a named range for the spilled cells, and uses that name in a SUM formula. | Explain how to obtain the CellArea of a dynamic array spill in Aspose.Cells and assign it to a workbook name. | Provide steps to reference a dynamic‑array spill via a named range in subsequent formulas and save the workbook using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Learn how to set a SEQUENCE dynamic‑array formula, refresh the spill, capture its CellArea, define a workbook named range that points to the spilled cells, and reference that name in a SUM (or other) formula using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        Cells cells = ws.Cells;

        // Populate some data that the dynamic array formula will depend on
        cells["B1"].PutValue(1);
        cells["B2"].PutValue(2);
        cells["B3"].PutValue(3);

        // Set a dynamic array formula in A1 that will spill based on B3 (e.g., SEQUENCE)
        Cell dynCell = cells["A1"];
        CellArea spillArea = dynCell.SetDynamicArrayFormula("=SEQUENCE(B3)", new FormulaParseOptions(), true);

        // Refresh dynamic array formulas so the spill range is materialized
        wb.RefreshDynamicArrayFormulas(true);

        // Build the address string for the spilled range (e.g., Sheet1!$A$1:$A$3)
        string spillAddress = $"={ws.Name}!{cells[spillArea.StartRow, spillArea.StartColumn].Name}:{cells[spillArea.EndRow, spillArea.EndColumn].Name}";

        // Create a named range that refers to the spilled range
        int nameIdx = wb.Worksheets.Names.Add("SpillRange");
        Name spillName = wb.Worksheets.Names[nameIdx];
        spillName.RefersTo = spillAddress;

        // Use the named range in another formula (e.g., sum of the spilled values)
        cells["D1"].Formula = "=SUM(SpillRange)";

        // Calculate all formulas
        wb.CalculateFormula();

        // Output the spilled values and the sum result
        Console.WriteLine("Spilled values:");
        for (int r = spillArea.StartRow; r <= spillArea.EndRow; r++)
        {
            for (int c = spillArea.StartColumn; c <= spillArea.EndColumn; c++)
            {
                Console.Write(cells[r, c].Value + "\t");
            }
            Console.WriteLine();
        }
        Console.WriteLine($"Sum of spilled range (D1): {cells["D1"].Value}");

        // Save the workbook
        wb.Save("DynamicArrayNamedRange.xlsx");
    }
}
