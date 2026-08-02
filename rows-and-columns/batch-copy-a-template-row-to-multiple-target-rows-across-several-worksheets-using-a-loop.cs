// Title: Batch copy a template row to multiple rows across worksheets with Aspose.Cells for .NET (C#)
// Description: This C# example creates a workbook, adds sample headers to the first row, generates additional worksheets, and uses a nested loop with Cells.CopyRow to duplicate the template row into several specified rows on every sheet, then saves the file as BatchCopyRowsOutput.xlsx.
// Keywords: Aspose.Cells | C# | CopyRow | batch copy rows | multiple worksheets | template row | Excel automation | Aspose.Cells API | copy row across sheets | GitHub example
// Common Searches: Aspose.Cells copy row to multiple sheets C# | How to duplicate a header row on every worksheet using Aspose.Cells | Batch copy rows across worksheets in .NET | Copy template row to specific rows with Aspose.Cells | Loop to copy rows in Excel using Aspose.Cells
// Developer Intent: Duplicate a single source row into several target rows on each worksheet of a workbook.
// Use Cases: Add identical header rows to all sheets in a financial report | Populate repeated product‑info rows in each section of multi‑sheet templates | Prepare a data‑entry template where the same row appears at multiple positions across worksheets
// AI Prompts: Show C# code that copies row 1 to rows 3, 6, and 9 on every worksheet using Aspose.Cells CopyRow. | Give an example that preserves cell styles, formats, and merged cells while batch copying a template row across multiple sheets. | Explain how to limit the copy operation to worksheets whose names begin with 'Data' using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace BatchCopyRowsExample
{
    // This C# example creates a workbook, adds sample headers to the first row, generates additional worksheets, and uses a nested loop with Cells.CopyRow to duplicate the template row into several specified rows on every sheet, then saves the file as BatchCopyRowsOutput.xlsx.
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();

            // ---------- Prepare the template row ----------
            // Use the first worksheet as the source of the template row (row index 0)
            Worksheet templateSheet = workbook.Worksheets[0];
            // Fill some sample data in the template row
            templateSheet.Cells["A1"].PutValue("Product");
            templateSheet.Cells["B1"].PutValue("Quantity");
            templateSheet.Cells["C1"].PutValue("Price");

            // ---------- Add additional worksheets ----------
            // For demonstration, create a total of 3 worksheets (including the first one)
            for (int i = 1; i < 3; i++)
            {
                workbook.Worksheets.Add($"Sheet{i + 1}");
            }

            // ---------- Define target rows where the template row will be copied ----------
            // These rows will receive a copy of the template row in each worksheet
            int[] targetRows = new int[] { 2, 5, 8 }; // zero‑based indices (row 3, 6, 9 in Excel)

            // ---------- Loop through each worksheet and copy the template row ----------
            foreach (Worksheet ws in workbook.Worksheets)
            {
                foreach (int targetRow in targetRows)
                {
                    // Copy the template row (row 0) from the template sheet to the target row
                    // in the current worksheet.
                    ws.Cells.CopyRow(templateSheet.Cells, 0, targetRow);
                }
            }

            // ---------- Save the result ----------
            workbook.Save("BatchCopyRowsOutput.xlsx");
        }
    }
}
