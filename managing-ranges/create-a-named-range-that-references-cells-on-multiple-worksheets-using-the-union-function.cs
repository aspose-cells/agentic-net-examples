// Title: Aspose.Cells C# – Create a Named Union Range Across Multiple Worksheets
// Description: Demonstrates how to build a UnionRange that spans cells on different worksheets using Workbook.Worksheets.CreateUnionRange, assign a name to the range, optionally set a common value, and save the workbook.
// Keywords: Aspose.Cells | C# | UnionRange | CreateUnionRange | named range | multiple worksheets | cross‑sheet range | Excel automation | .NET
// Common Searches: Aspose.Cells create union range across sheets | named range spanning multiple worksheets C# | CreateUnionRange example Aspose.Cells | set value for union range Aspose.Cells | reference cells from different sheets in a single range Aspose.Cells
// Developer Intent: Create a union range that references cells on several worksheets, give it a name, and optionally assign a value.
// Use Cases: Define a single named range that aggregates data from Sheet1!A1:A5 and Sheet2!B1:B5 for formulas or data validation. | Apply uniform formatting or a constant value to non‑contiguous blocks located on different worksheets. | Use the named union range as a source for chart series, pivot tables, or custom calculations across sheets.
// AI Prompts: Generate C# code that creates a UnionRange covering three worksheets, names it, and adds a SUM formula that totals all cells. | Explain how to retrieve the address string of a UnionRange created with CreateUnionRange and iterate through its individual areas. | Show how to apply conditional formatting to a UnionRange that spans multiple sheets using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsUnionRangeExample
{
    // Demonstrates how to build a UnionRange that spans cells on different worksheets using Workbook.Worksheets.CreateUnionRange, assign a name to the range, optionally set a common value, and save the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook with two worksheets
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";
            Worksheet sheet2 = workbook.Worksheets[workbook.Worksheets.Add()];
            sheet2.Name = "Sheet2";

            // Populate some data in the first worksheet (A1:A5)
            for (int i = 0; i < 5; i++)
                sheet1.Cells[i, 0].PutValue($"S1_R{i + 1}");

            // Populate some data in the second worksheet (B1:B5)
            for (int i = 0; i < 5; i++)
                sheet2.Cells[i, 1].PutValue($"S2_R{i + 1}");

            // Create a union range that spans both worksheets.
            // The address string includes sheet names and the individual ranges,
            // separated by commas. The sheetIndex parameter is the index of the
            // first sheet referenced (Sheet1 = 0).
            UnionRange unionRange = workbook.Worksheets.CreateUnionRange(
                "Sheet1!A1:A5,Sheet2!B1:B5", 0);

            // Assign a name to the union range. This name can be used in formulas.
            unionRange.Name = "MyUnionRange";

            // Optionally set a value for the entire union range to visualize it.
            unionRange.Value = "Combined";

            // Save the workbook.
            workbook.Save("UnionRangeMultipleSheets.xlsx");
        }
    }
}
