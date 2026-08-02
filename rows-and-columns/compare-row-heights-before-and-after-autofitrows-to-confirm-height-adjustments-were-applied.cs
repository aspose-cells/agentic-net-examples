// Title: Compare Row Heights Before and After AutoFitRows with Aspose.Cells for .NET
// Description: Shows how to capture a row's height, run Worksheet.AutoFitRows(), retrieve the new height, and confirm that the height was adjusted when wrapped text expands the row.
// Keywords: Aspose.Cells AutoFitRows | C# get row height | compare row heights | detect row height change | worksheet autofit rows .NET
// Common Searches: How to check if AutoFitRows changed row height in Aspose.Cells | Get row height before and after AutoFitRows C# | Verify row height adjustment after Worksheet.AutoFitRows | Aspose.Cells row height comparison example
// Developer Intent: Find out whether calling AutoFitRows modifies the height of a specific row.
// Use Cases: Validate that text wrapping triggers a taller row after AutoFitRows and log the before/after values. | Create automated tests that assert row‑height changes when generating reports. | Conditionally apply extra formatting only when AutoFitRows increases a row's height.
// AI Prompts: Generate C# code using Aspose.Cells to record a row's height, call AutoFitRows, and output the height difference. | Provide a method that returns true if any row height was altered by AutoFitRows in a workbook. | Show how to assert programmatically that AutoFitRows increased the height of a wrapped‑text row.

using System;
using Aspose.Cells;

namespace RowHeightComparisonDemo
{
    // Shows how to capture a row's height, run Worksheet.AutoFitRows(), retrieve the new height, and confirm that the height was adjusted when wrapped text expands the row.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate cells with long text and enable text wrapping to force row height changes
            cells["A1"].PutValue("This is a very long piece of text that should cause the row height to increase after AutoFitRows is applied.");
            Style wrapStyle = cells["A1"].GetStyle();
            wrapStyle.IsTextWrapped = true;
            cells["A1"].SetStyle(wrapStyle);

            // Optionally add more data in the same row to demonstrate effect
            cells["B1"].PutValue("Additional content in the same row.");

            // Capture the row height before AutoFitRows
            double heightBefore = cells.GetRowHeight(0);
            Console.WriteLine($"Row 0 height before AutoFitRows: {heightBefore}");

            // AutoFit all rows in the worksheet
            worksheet.AutoFitRows();

            // Capture the row height after AutoFitRows
            double heightAfter = cells.GetRowHeight(0);
            Console.WriteLine($"Row 0 height after AutoFitRows: {heightAfter}");

            // Compare heights and confirm adjustment
            if (Math.Abs(heightBefore - heightAfter) > 0.01)
            {
                Console.WriteLine("Row height was adjusted by AutoFitRows.");
            }
            else
            {
                Console.WriteLine("Row height remained unchanged.");
            }

            // Save the workbook to verify the result manually if needed
            workbook.Save("RowHeightComparisonDemo.xlsx");
        }
    }
}
