// Title: Automatically Update a Named Range After Inserting Rows and Refresh Dynamic Array Formulas with Aspose.Cells for .NET
// Description: Demonstrates how to create a static named range, insert new rows, extend the range by updating its RefersTo property, and call RefreshDynamicArrayFormulas so that dependent dynamic‑array formulas spill correctly. The example uses C# and Aspose.Cells.
// Keywords: Aspose.Cells C# named range update | refresh dynamic array formulas Aspose.Cells | extend named range after inserting rows | RefersTo property Aspose.Cells | dynamic array spill range .NET | programmatic range expansion | Excel automation Aspose.Cells
// Common Searches: how to extend a named range after adding rows in Aspose.Cells | refresh dynamic array formulas after inserting rows C# | change RefersTo for a named range programmatically Aspose.Cells | update static named range automatically | Aspose.Cells insert rows keep named range current
// Developer Intent: Keep a named range synchronized with added rows and ensure any dynamic‑array formulas that reference it expand automatically.
// Use Cases: Create a static named range, insert rows, modify the RefersTo address, then refresh dynamic array formulas to update the spill area. | Replace the static range with a dynamic OFFSET or Table formula so the range adjusts without code changes. | Generate workbooks where data length varies, inserting rows on the fly while maintaining accurate named ranges and dependent formulas.
// AI Prompts: Generate C# code using Aspose.Cells that automatically adjusts a named range when rows are inserted, eliminating manual RefersTo updates. | Explain when and why RefreshDynamicArrayFormulas should be called after modifying worksheet data in Aspose.Cells. | Show how to define a dynamic named range with OFFSET or a Table in Aspose.Cells so formulas stay up‑to‑date without extra logic.

using System;
using Aspose.Cells;

namespace AsposeCellsDynamicNamedRangeDemo
{
    // Demonstrates how to create a static named range, insert new rows, extend the range by updating its RefersTo property, and call RefreshDynamicArrayFormulas so that dependent dynamic‑array formulas spill correctly. The example uses C# and Aspose.Cells.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate initial data in column A (A1:A5)
            for (int i = 0; i < 5; i++)
            {
                cells[i, 0].PutValue(i + 1); // Values 1..5
            }

            // Create a named range "MyData" that refers to the current data range A1:A5
            int nameIndex = workbook.Worksheets.Names.Add("MyData");
            Name myDataName = workbook.Worksheets.Names[nameIndex];
            myDataName.RefersTo = $"=Sheet1!$A$1:$A$5";

            // Set a dynamic array formula in cell B1 that spills the values of the named range
            // The "#" operator returns the spill range of the dynamic array
            cells["B1"].SetDynamicArrayFormula("=MyData#", new FormulaParseOptions(), true);

            // Insert two new rows after the existing data (at row index 5, i.e., after A5)
            // This will shift existing rows down and expand the data area
            cells.InsertRows(5, 2, true);

            // Populate the newly inserted rows with additional data
            cells[5, 0].PutValue(6);
            cells[6, 0].PutValue(7);

            // Update the named range to include the new rows.
            // Since the name refers to a static range, we need to adjust its RefersTo property.
            // Alternatively, if the name used a dynamic formula (e.g., OFFSET), this step could be omitted.
            myDataName.RefersTo = $"=Sheet1!$A$1:$A$7";

            // Refresh dynamic array formulas so that the spill range in B1 expands automatically
            workbook.RefreshDynamicArrayFormulas(true);

            // Save the workbook to verify the result
            workbook.Save("DynamicNamedRangeUpdated.xlsx");
        }
    }
}
