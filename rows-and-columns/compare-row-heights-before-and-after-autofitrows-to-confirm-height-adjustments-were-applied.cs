// Title: How to Compare Row Height Before and After AutoFitRows in Aspose.Cells for .NET
// Description: A C# example that creates a workbook, writes wrapped text to a cell, records the row height, runs worksheet.AutoFitRows(), reads the new height, and verifies that the row height was adjusted before saving the file.
// Keywords: Aspose.Cells | .NET | C# | AutoFitRows | row height | get row height | compare row heights | wrapped text | worksheet.AutoFitRows | row height adjustment
// Common Searches: Aspose.Cells check row height after AutoFitRows | C# get row height before and after AutoFitRows | verify AutoFitRows changes row height | how to measure row height in Aspose.Cells | compare row heights in .NET spreadsheet
// Developer Intent: Determine whether calling AutoFitRows modifies a row's height by comparing the measured values before and after the operation.
// Use Cases: Validate that wrapped text triggers a height increase when AutoFitRows is applied to dynamic reports. | Log before‑and‑after row heights to audit formatting changes during automated spreadsheet generation. | Create unit tests that assert AutoFitRows correctly expands rows containing multi‑line content.
// AI Prompts: Generate a C# function using Aspose.Cells that returns true if any row height changes after worksheet.AutoFitRows(). | Write code that captures heights of all rows, runs AutoFitRows, and prints rows where the height difference exceeds 0.01. | Provide a unit‑test example that confirms AutoFitRows increases the height of a cell with wrapped text.

using System;
using Aspose.Cells;

namespace RowHeightComparisonDemo
{
    // A C# example that creates a workbook, writes wrapped text to a cell, records the row height, runs worksheet.AutoFitRows(), reads the new height, and verifies that the row height was adjusted before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate a cell with long wrapped text to force a height change
            cells["A1"].PutValue("This is a very long text that should cause the row height to increase after auto‑fit.");
            Style style = cells["A1"].GetStyle();
            style.IsTextWrapped = true;               // Enable text wrapping
            cells["A1"].SetStyle(style);

            // Optionally set a fixed column width to make wrapping effective
            cells.SetColumnWidth(0, 20);

            // Retrieve the initial row height (before auto‑fit)
            double heightBefore = cells.GetRowHeight(0);
            Console.WriteLine($"Row 0 height before AutoFitRows: {heightBefore}");

            // Perform auto‑fit on all rows in the worksheet
            worksheet.AutoFitRows();

            // Retrieve the row height after auto‑fit
            double heightAfter = cells.GetRowHeight(0);
            Console.WriteLine($"Row 0 height after AutoFitRows: {heightAfter}");

            // Compare the heights and confirm adjustment
            if (Math.Abs(heightAfter - heightBefore) > 0.01)
            {
                Console.WriteLine("Row height was adjusted by AutoFitRows.");
            }
            else
            {
                Console.WriteLine("Row height remained unchanged.");
            }

            // Save the workbook (lifecycle rule)
            workbook.Save("RowHeightComparisonDemo.xlsx");
        }
    }
}
