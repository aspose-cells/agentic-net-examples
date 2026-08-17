// Title: Aspose.Cells .NET – Retrieve spilled range of a dynamic array formula in C3
// Description: C# example that creates a workbook, inserts a dynamic array formula (e.g., =SEQUENCE(3,2)) into cell C3, forces calculation, and uses GetArrayRange together with CellsHelper.CellIndexToName to return the full A1‑style spill address.
// Keywords: Aspose.Cells GetArrayRange | dynamic array spill address | C# Aspose.Cells dynamic array | SEQUENCE formula spill range | CellsHelper CellIndexToName | retrieve spilled range .NET | Excel dynamic array Aspose
// Common Searches: how to get spilled range of a dynamic array formula using Aspose.Cells | Aspose.Cells GetArrayRange example C# | retrieve A1 address of SEQUENCE spill in C3 | Aspose.Cells dynamic array spill range code
// Developer Intent: Obtain the A1‑style address of the range that a dynamic array formula occupies when placed in cell C3.
// Use Cases: Apply formatting or borders to the exact area produced by a SEQUENCE formula. | Validate that a dynamic array output stays within worksheet limits before saving. | Reference the spill range in subsequent calculations or data‑processing logic.
// AI Prompts: Generate C# code with Aspose.Cells that returns the spilled range address of a dynamic array formula in cell C3. | Explain how GetArrayRange and CellsHelper.CellIndexToName combine to produce an A1‑style spill address. | Show how to change the formula to =SORT(A1:A10) and retrieve its spill range using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsDynamicArraySpill
{
    // C# example that creates a workbook, inserts a dynamic array formula (e.g., =SEQUENCE(3,2)) into cell C3, forces calculation, and uses GetArrayRange together with CellsHelper.CellIndexToName to return the full A1‑style spill address.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule)
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set a dynamic array formula in cell C3 (example: SEQUENCE(3,2))
            Cell targetCell = cells["C3"];
            string formula = "=SEQUENCE(3,2)";
            targetCell.SetDynamicArrayFormula(formula, new FormulaParseOptions(), true);

            // Calculate formulas so the spill range is materialized
            wb.CalculateFormula();

            // Retrieve the spilled range of the dynamic array formula
            CellArea spillArea = targetCell.GetArrayRange();

            // Convert the start and end coordinates to A1 style addresses
            string startAddress = CellsHelper.CellIndexToName(spillArea.StartRow, spillArea.StartColumn);
            string endAddress   = CellsHelper.CellIndexToName(spillArea.EndRow,   spillArea.EndColumn);
            string spilledRange = $"{startAddress}:{endAddress}";

            // Output the spilled range address
            Console.WriteLine($"Spilled range for dynamic array formula in C3: {spilledRange}");

            // (Optional) Save the workbook to verify the result
            wb.Save("DynamicArraySpillDemo.xlsx");
        }
    }
}
