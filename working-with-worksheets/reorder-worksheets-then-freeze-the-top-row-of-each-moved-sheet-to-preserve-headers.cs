// Title: C# Aspose.Cells: Reorder Worksheets and Freeze Header Row in Every Sheet
// Description: Shows how to create a workbook, rename and add sheets, move a worksheet to a new index with MoveTo, loop through all worksheets to freeze the first row using FreezePanes, and save the result.
// Keywords: Aspose.Cells | C# | .NET | reorder worksheets | move worksheet position | FreezePanes | freeze header row | Excel sheet ordering | programmatic Excel automation | Aspose.Cells workbook manipulation
// Common Searches: Aspose.Cells move worksheet to first position C# | How to freeze top row on every sheet using Aspose.Cells | Reorder Excel sheets programmatically .NET | Freeze panes for all worksheets Aspose.Cells | C# code to change sheet order and freeze header row
// Developer Intent: Programmatically change the order of worksheets in an Excel workbook and apply a frozen top row to each sheet so column headers stay visible while scrolling.
// Use Cases: Generate a multi‑sheet report where the summary sheet is placed at the front and every sheet’s header row remains visible. | Create an export routine that adds new worksheets, moves them to a specific position, and freezes the first row for easier data review. | Prepare a user‑friendly workbook for distribution, ensuring logical sheet order and locked header rows across all tabs.
// AI Prompts: Write C# code with Aspose.Cells to move a worksheet named "Data" to index 2 and freeze the first two rows of every sheet. | Provide an example that sorts all worksheets alphabetically and then applies FreezePanes to lock the header row on each worksheet. | Show how to loop through a workbook's worksheets, move a specific sheet to the first position, and freeze the top row for each sheet before saving.

using System;
using Aspose.Cells;

namespace AsposeCellsReorderAndFreeze
{
    // Shows how to create a workbook, rename and add sheets, move a worksheet to a new index with MoveTo, loop through all worksheets to freeze the first row using FreezePanes, and save the result.
    class Program
    {
        static void Main()
        {
            // Create a new workbook with the default first worksheet
            Workbook workbook = new Workbook();

            // Rename the default sheet and add two more sheets
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";

            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");

            // Example re‑ordering: move "Sheet3" to the first position (index 0)
            // This shifts the existing sheets to the right.
            sheet3.MoveTo(0);

            // Freeze the top row (row 1) of every worksheet to keep headers visible
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Freeze at row index 1 (second row), column index 0 (first column)
                // with 1 frozen row and 0 frozen columns.
                ws.FreezePanes(1, 0, 1, 0);
            }

            // Save the modified workbook
            workbook.Save("ReorderedAndFrozen.xlsx");
        }
    }
}
