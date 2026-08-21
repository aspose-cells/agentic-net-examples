// Title: Refresh a FILTER dynamic array after source data change with Aspose.Cells for .NET
// Description: This C# example demonstrates how to create a workbook, populate source columns, apply a FILTER dynamic‑array formula, modify the criteria values, and invoke RefreshDynamicArrayFormulas(true) to recalculate and spill the updated results. The workbook is then saved with the refreshed data.
// Keywords: Aspose.Cells .NET | FILTER dynamic array | RefreshDynamicArrayFormulas | recalculate spilled range | update source data | dynamic array formula refresh | C# spreadsheet automation | Excel FILTER function programmatic
// Common Searches: Aspose.Cells refresh FILTER formula after data change | C# recalculate dynamic array formulas | how to update spilled range in Aspose.Cells | RefreshDynamicArrayFormulas usage example | dynamic array formula recalc .NET
// Developer Intent: Update source data for a FILTER dynamic array and recalculate the spilled results using Aspose.Cells.
// Use Cases: Create a new workbook and fill columns with source and criteria values. | Set a FILTER dynamic‑array formula that spills into adjacent cells. | Modify the criteria range to change which rows meet the filter condition. | Call RefreshDynamicArrayFormulas(true) to automatically refresh the spill area. | Print the refreshed values and save the workbook.
// AI Prompts: Write C# code that changes the criteria range of a FILTER dynamic array and uses Aspose.Cells to refresh the spilled results. | Explain how RefreshDynamicArrayFormulas(true) works in Aspose.Cells and why it is needed after modifying source data. | Provide a step‑by‑step tutorial for setting, updating, and saving a FILTER dynamic array in a .NET workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsDynamicArrayRefreshDemo
{
    // This C# example demonstrates how to create a workbook, populate source columns, apply a FILTER dynamic‑array formula, modify the criteria values, and invoke RefreshDynamicArrayFormulas(true) to recalculate and spill the updated results. The workbook is then saved with the refreshed data.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate source data for the FILTER formula
                // Column A: values to be filtered
                cells["A2"].PutValue(10);
                cells["A3"].PutValue(20);
                cells["A4"].PutValue(30);
                cells["A5"].PutValue(40);

                // Column B: criteria values
                cells["B2"].PutValue(40);
                cells["B3"].PutValue(60);
                cells["B4"].PutValue(55);
                cells["B5"].PutValue(30);

                // Set a FILTER dynamic array formula in C1
                // It will spill results into C1:Cn depending on the criteria
                cells["C1"].SetDynamicArrayFormula(
                    "=FILTER(A2:A5, B2:B5>50)",
                    new FormulaParseOptions(),
                    true); // calculateValue = true

                // Calculate initial formulas (optional, RefreshDynamicArrayFormulas can also calculate)
                workbook.CalculateFormula();

                Console.WriteLine("Initial FILTER results:");
                PrintSpillRange(cells, startRow: 0, startColumn: 2); // C column (index 2)

                // Change the source data that influences the FILTER formula
                cells["B2"].PutValue(70); // Now this row meets the >50 condition
                cells["B4"].PutValue(80); // This row also meets the condition

                // Refresh dynamic array formulas and recalculate their values
                workbook.RefreshDynamicArrayFormulas(true);

                Console.WriteLine("\nAfter data change and refresh:");
                PrintSpillRange(cells, startRow: 0, startColumn: 2);

                // Save the workbook (lifecycle: save)
                workbook.Save("DynamicArrayRefreshResult.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Helper method to print the spilled range of a dynamic array formula starting at a given cell
        static void PrintSpillRange(Cells cells, int startRow, int startColumn)
        {
            int row = startRow;
            while (true)
            {
                Cell cell = cells[row, startColumn];
                // Stop when the cell is null or contains no value
                if (cell == null || cell.Value == null)
                    break;

                Console.WriteLine($"Cell {cell.Name}: {cell.Value}");
                row++;
            }
        }
    }
}
