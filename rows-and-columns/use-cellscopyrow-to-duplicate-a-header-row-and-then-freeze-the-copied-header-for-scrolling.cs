// Title: How to duplicate the first header row and freeze the top two rows in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to copy row 0 to row 1 with Cells.CopyRow and then apply Worksheet.FreezePanes to keep rows 0‑1 visible while scrolling. | Show an example of freezing the first two rows after copying a header row in an Excel file using Aspose.Cells for .NET. | Provide a step‑by‑step C# snippet that creates a workbook, adds a header, duplicates it to the next row, and freezes both rows with Aspose.Cells.
// Common Searches: Aspose.Cells C# copy header row and freeze top rows | How to use Cells.CopyRow to duplicate a row then freeze panes in Aspose.Cells | C# Aspose.Cells freeze first two rows after copying a row | Duplicate Excel header and keep it visible while scrolling with Aspose.Cells | Freeze panes for copied header row using Aspose.Cells .NET
// Tags: Cells.CopyRow duplicate header C# | Worksheet.FreezePanes top rows .NET | Aspose.Cells copy row and freeze panes | Excel header duplication Aspose.Cells | freeze first two rows Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsHeaderCopyAndFreeze
{
    // The example creates a new workbook, writes a header in the first row, adds sample data, copies the header to the second row using Cells.CopyRow, freezes the first two rows with Worksheet.FreezePanes, and saves the file as HeaderCopyAndFreeze.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ----- Populate original header row (row 0) -----
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Name");
            cells["C1"].PutValue("Date");

            // Add some sample data rows
            for (int i = 1; i <= 10; i++)
            {
                cells[i, 0].PutValue(i);                     // ID
                cells[i, 1].PutValue($"Item {i}");           // Name
                cells[i, 2].PutValue(DateTime.Today.AddDays(i)); // Date
            }

            // ----- Duplicate the header row to row 1 (second row) -----
            // CopyRow(sourceCells, sourceRowIndex, destinationRowIndex)
            cells.CopyRow(cells, 0, 1);

            // ----- Freeze panes so that the copied header (rows 0 and 1) stay visible -----
            // FreezePanes(row, column, freezedRows, freezedColumns)
            // Row index is zero‑based; to freeze first two rows we set row = 2 and freezedRows = 2
            sheet.FreezePanes(2, 0, 2, 0);

            // Save the workbook (save rule)
            workbook.Save("HeaderCopyAndFreeze.xlsx");
        }
    }
}
