// Title: C# AspNet: Set a Dynamic Array Formula that Spills and Sum It on a Different Sheet with Aspose.Cells
// Description: This example shows how to create a workbook, fill cells B1:B5 on Sheet1, assign a dynamic array formula "=B1:B5" to A1 so it spills into A1:A5, add Sheet2, reference the spilled range using the # operator in a SUM formula ("=SUM(Sheet1!A1#)"), refresh dynamic arrays, calculate all formulas, and retrieve the total value. The code also demonstrates saving the file as DynamicArraySpillSum.xlsx.
// Keywords: Aspose.Cells | C# | .NET | dynamic array formula | spill range | # operator | SUM across worksheets | RefreshDynamicArrayFormulas | CalculateFormula | Excel automation
// Common Searches: Aspose.Cells set dynamic array formula C# | reference spilled range on another sheet Aspose.Cells | sum spilled array values using # operator | refresh dynamic arrays Aspose.Cells .NET | calculate workbook formulas after spill
// Developer Intent: Create a spilling dynamic array in one sheet and use its # reference in a SUM formula on a second sheet.
// Use Cases: Automatically expand a formula to a variable‑size range without hard‑coding the size. | Consume a spilled array from a different worksheet for aggregation or further calculations. | Ensure spilled data is up‑to‑date by refreshing dynamic arrays before performing workbook calculations.
// AI Prompts: Generate C# code with Aspose.Cells that sets "=B1:B5" as a dynamic array in Sheet1!A1 and sums the spill on Sheet2 using "=SUM(Sheet1!A1#)". | Explain the purpose of the # operator when referencing a spilled range in Aspose.Cells and how to refresh dynamic arrays before calculation. | Provide step‑by‑step instructions to build a workbook, populate source data, apply a dynamic array formula, reference the spill on another sheet, refresh, calculate, and output the sum.

using System;
using Aspose.Cells;

namespace AsposeCellsDynamicArraySpillExample
{
    // This example shows how to create a workbook, fill cells B1:B5 on Sheet1, assign a dynamic array formula "=B1:B5" to A1 so it spills into A1:A5, add Sheet2, reference the spilled range using the # operator in a SUM formula ("=SUM(Sheet1!A1#)"), refresh dynamic arrays, calculate all formulas, and retrieve the total value. The code also demonstrates saving the file as DynamicArraySpillSum.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Sheet1: set up data and a dynamic array formula
            // -------------------------------------------------
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";
            Cells cells1 = sheet1.Cells;

            // Populate source data in B1:B5 (values 1 to 5)
            for (int i = 0; i < 5; i++)
            {
                cells1[i, 1].PutValue(i + 1); // Column index 1 = B
            }

            // Set a dynamic array formula in A1 that spills the range B1:B5
            // The formula "=B1:B5" will spill into A1:A5
            Cell dynamicCell = cells1["A1"];
            dynamicCell.SetDynamicArrayFormula("=B1:B5", new FormulaParseOptions(), true);

            // -------------------------------------------------
            // Sheet2: use the spill range in a SUM formula
            // -------------------------------------------------
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            Cells cells2 = sheet2.Cells;

            // Reference the spill range from Sheet1 using the # operator
            // "=SUM(Sheet1!A1#)" sums all values spilled from the dynamic array in Sheet1!A1
            Cell sumCell = cells2["A1"];
            sumCell.Formula = "=SUM(Sheet1!A1#)";

            // -------------------------------------------------
            // Refresh dynamic arrays and calculate all formulas
            // -------------------------------------------------
            // Refresh spill ranges (calculate = true to also compute values)
            workbook.RefreshDynamicArrayFormulas(true);

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // -------------------------------------------------
            // Output the result from Sheet2!A1
            // -------------------------------------------------
            Console.WriteLine("Sum of the spilled range from Sheet1: " + sumCell.Value);

            // -------------------------------------------------
            // Save the workbook (optional, demonstrates lifecycle rule)
            // -------------------------------------------------
            workbook.Save("DynamicArraySpillSum.xlsx");
        }
    }
}
