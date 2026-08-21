// Title: Merge Cells Across Hidden Rows and Verify Visibility with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to hide specific rows, merge a cell range that spans those hidden rows, and programmatically confirm that the rows stay hidden after the merge using Aspose.Cells for .NET. The workbook is saved as an example output.
// Keywords: Aspose.Cells merge hidden rows | C# merge cells across hidden rows | preserve row visibility after merge | verify hidden rows Aspose.Cells | Aspose.Cells hide rows then merge | Excel merge range hidden rows .NET
// Common Searches: merge cells that include hidden rows Aspose.Cells | does merging keep hidden rows hidden in .NET | check row hidden status after merge C# | Aspose.Cells hide rows and merge range example | how to preserve hidden rows when merging cells
// Developer Intent: The developer needs to merge a range of cells that contains hidden rows and ensure those rows remain hidden after the operation.
// Use Cases: Create a report header that spans multiple rows, some of which are hidden for layout, while keeping the hidden rows concealed. | Build a spreadsheet template that merges cells across hidden rows to group data without altering row visibility. | Automate Excel generation and validate that row visibility is unchanged after merging cells for UI consistency.
// AI Prompts: Generate C# code using Aspose.Cells to merge a range that includes hidden rows and then verify the rows stay hidden. | Explain Aspose.Cells' behavior with hidden rows during a merge and show how to check their hidden status afterward. | Suggest alternative techniques to merge cells without affecting the hidden property of rows in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace MergeHiddenRowsDemo
{
    // Demonstrates how to hide specific rows, merge a cell range that spans those hidden rows, and programmatically confirm that the rows stay hidden after the merge using Aspose.Cells for .NET. The workbook is saved as an example output.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some sample data in rows 0 to 6, columns 0 to 2
            for (int row = 0; row <= 6; row++)
            {
                for (int col = 0; col <= 2; col++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Hide rows 2, 3, and 4 (zero‑based indices 2‑4)
            cells.HideRows(2, 3); // hides rows 2,3,4

            // Merge a range that includes the hidden rows:
            // From row 1 (second row) column 0 to row 5 (sixth row) column 2
            // This range covers rows 1‑5, i.e., includes the hidden rows 2‑4
            cells.Merge(1, 0, 5, 3); // totalRows = 5, totalColumns = 3

            // Verify that the hidden rows are still hidden after merging
            for (int row = 2; row <= 4; row++)
            {
                bool isHidden = worksheet.Cells.Rows[row].IsHidden;
                Console.WriteLine($"Row {row + 1} hidden status after merge: {isHidden}");
            }

            // Save the workbook to demonstrate the result
            workbook.Save("MergeHiddenRowsDemo.xlsx");
        }
    }
}
