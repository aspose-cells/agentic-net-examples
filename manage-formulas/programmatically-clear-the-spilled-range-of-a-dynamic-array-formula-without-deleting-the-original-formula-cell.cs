// Title: Aspose.Cells .NET – Clear Dynamic Array Spill Range While Preserving Source Formula
// Description: Demonstrates how to set a dynamic array formula (e.g., =SEQUENCE(5)) in a workbook, retrieve its full spill area with GetArrayRange, clear only the spilled cells using ClearRange, optionally refresh formulas with RefreshDynamicArrayFormulas, and save the file—leaving the original formula cell untouched.
// Keywords: Aspose.Cells | .NET | dynamic array | spill range | ClearRange | GetArrayRange | SetDynamicArrayFormula | RefreshDynamicArrayFormulas | C# example | Excel dynamic array | global | US developers | EU developers
// Common Searches: How to clear a dynamic array spill in Aspose.Cells C# | Aspose.Cells GetArrayRange usage | Remove spilled cells without deleting formula Aspose | Refresh dynamic array formulas after clearing spill .NET | C# code to reset SEQUENCE spill range | Aspose.Cells tutorial for dynamic array management
// Developer Intent: Delete only the spilled cells of a dynamic array formula while keeping the formula cell intact.
// Use Cases: Clear previous results of a SEQUENCE formula before applying new parameters in automated reporting. | Reset spilled data after external data source changes without recreating the formula. | Maintain formula integrity while cleaning up stale spill values during batch workbook processing. | Prepare a template workbook for reuse by wiping spill content while preserving formulas.
// AI Prompts: Generate C# code using Aspose.Cells that obtains the spill area of a dynamic array formula and clears it without affecting the source cell. | Explain the role of GetArrayRange and ClearRange for resetting dynamic array results in a .NET workbook. | Show how to call RefreshDynamicArrayFormulas after clearing a spill to keep the workbook consistent. | Provide step‑by‑step instructions for preserving a dynamic array formula while deleting its spilled values.

using System;
using Aspose.Cells;

namespace AsposeCellsDynamicArrayClearSpill
{
    // Demonstrates how to set a dynamic array formula (e.g., =SEQUENCE(5)) in a workbook, retrieve its full spill area with GetArrayRange, clear only the spilled cells using ClearRange, optionally refresh formulas with RefreshDynamicArrayFormulas, and save the file—leaving the original formula cell untouched.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Set a dynamic array formula in cell A1 (will spill to A1:A5)
                Cell formulaCell = cells["A1"];
                string formula = "=SEQUENCE(5)";
                formulaCell.SetDynamicArrayFormula(formula, new FormulaParseOptions(), true);

                // Calculate formulas so the spill range is populated
                workbook.CalculateFormula();

                // Retrieve the full spilled range of the dynamic array formula
                CellArea spillArea = formulaCell.GetArrayRange();

                // Clear the entire spill range
                cells.ClearRange(spillArea);

                // Re‑apply the original dynamic array formula to the original cell
                formulaCell.SetDynamicArrayFormula(formula, new FormulaParseOptions(), true);

                // Re‑calculate to populate the spill range again
                workbook.CalculateFormula();

                // Refresh dynamic array formulas (optional, ensures consistency)
                workbook.RefreshDynamicArrayFormulas(true);

                // Save the workbook
                workbook.Save("DynamicArraySpillCleared.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}
