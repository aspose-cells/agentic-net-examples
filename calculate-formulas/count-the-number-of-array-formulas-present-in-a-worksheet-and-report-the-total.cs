// Title: Count legacy and dynamic array formulas in an Aspose.Cells worksheet (C#)
// Description: This example creates a workbook, inserts a CSE‑style array formula and a dynamic array formula, forces calculation so the spill range materializes, then scans all used cells using the IsArrayFormula and IsDynamicArrayFormula flags to compute the total number of array‑formula cells and outputs the count.
// Keywords: Aspose.Cells array formula count | C# IsArrayFormula | IsDynamicArrayFormula | legacy CSE array | dynamic spill formula | Excel audit .NET | spreadsheet automation | count array cells
// Common Searches: Aspose.Cells count array formulas C# | detect dynamic array formulas with Aspose.Cells | enumerate legacy array formulas in Excel workbook C# | total array formula cells Aspose.Cells .NET | list spilled array cells Aspose.Cells
// Developer Intent: Determine how many cells belong to any array formula—legacy or dynamic—in a worksheet.
// Use Cases: Validate that a generated report contains the expected number of array calculations before distribution. | Create an audit report summarizing array‑formula usage across all worksheets in a large workbook. | Skip array‑formula cells during custom data‑processing loops. | Log the presence of spilled dynamic arrays prior to exporting data to CSV.
// AI Prompts: Provide a C# function that returns a dictionary mapping each worksheet name to its array‑formula cell count using Aspose.Cells. | Show code to collect the addresses of every cell that participates in a dynamic array spill. | Explain the difference between IsArrayFormula and IsDynamicArrayFormula properties and when to use each.

using System;
using Aspose.Cells;

// This example creates a workbook, inserts a CSE‑style array formula and a dynamic array formula, forces calculation so the spill range materializes, then scans all used cells using the IsArrayFormula and IsDynamicArrayFormula flags to compute the total number of array‑formula cells and outputs the count.
class CountArrayFormulas
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate some sample data
        cells["A1"].PutValue(1);
        cells["A2"].PutValue(2);
        cells["A3"].PutValue(3);

        // Add a legacy (CSE) array formula that spans 3 rows and 1 column
        cells["B1"].SetArrayFormula("=A1:A3*2", 3, 1);

        // Add a dynamic array formula (will spill into neighboring cells)
        cells["D1"].SetDynamicArrayFormula("=SEQUENCE(4)", new FormulaParseOptions(), true);

        // Calculate formulas so that the dynamic array spills are materialized
        workbook.CalculateFormula();

        // Count cells that contain either a legacy array formula or a dynamic array formula
        int arrayFormulaCount = 0;
        int maxRow = cells.MaxDataRow;
        int maxCol = cells.MaxDataColumn;

        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = cells[row, col];
                if (cell.IsArrayFormula || cell.IsDynamicArrayFormula)
                {
                    arrayFormulaCount++;
                }
            }
        }

        Console.WriteLine("Total array formulas in the worksheet: " + arrayFormulaCount);

        // Save the workbook (optional)
        workbook.Save("ArrayFormulaCountDemo.xlsx");
    }
}
