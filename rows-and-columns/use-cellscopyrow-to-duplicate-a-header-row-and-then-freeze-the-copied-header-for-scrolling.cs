// Title: Copy Header Row and Freeze Top Two Rows with Cells.CopyRow & FreezePanes (C# Aspose.Cells)
// Description: C# example that creates a workbook, writes a header in row 0, duplicates it to row 1 using Cells.CopyRow, then freezes rows 0‑1 with Worksheet.FreezePanes so the headers stay visible while scrolling, and saves the file.
// Keywords: Aspose.Cells C# copy row | Cells.CopyRow example | FreezePanes C# | duplicate header Aspose.Cells | freeze top rows Excel | Aspose.Cells worksheet freeze | C# Excel automation | Aspose.Cells GitHub sample | Excel export header repeat | Aspose.Cells API
// Common Searches: Aspose.Cells copy row and freeze panes C# | How to duplicate a header row in Aspose.Cells | Freeze first two rows after copying header Aspose.Cells | Cells.CopyRow usage example | Worksheet.FreezePanes parameters C# | Aspose.Cells header repeat for printing
// Developer Intent: Copy the first worksheet row to a second row and keep both rows frozen for constant visibility during scrolling.
// Use Cases: Generate reports where a repeated header is needed on the second row for sub‑section titles while keeping both rows static. | Create printable Excel exports with a duplicated header row that remains visible on screen and on printed pages. | Develop data‑grid exports where the top two rows are frozen after copying the header to preserve context during navigation.
// AI Prompts: Show a C# snippet that uses Cells.CopyRow to duplicate row 0 to row 1 and then freezes rows 0‑1 with FreezePanes in Aspose.Cells. | Explain the FreezePanes parameters for freezing multiple rows after copying a header with Cells.CopyRow. | Provide a complete Aspose.Cells example that copies a header row, freezes the top two rows, and saves the workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsHeaderCopyAndFreeze
{
    // C# example that creates a workbook, writes a header in row 0, duplicates it to row 1 using Cells.CopyRow, then freezes rows 0‑1 with Worksheet.FreezePanes so the headers stay visible while scrolling, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ----- Populate original header row (row 0) -----
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Name");
            cells["C1"].PutValue("Date");

            // Add some sample data rows
            for (int i = 2; i <= 10; i++)
            {
                cells[i - 1, 0].PutValue(i - 1);                     // ID
                cells[i - 1, 1].PutValue($"Item {i - 1}");          // Name
                cells[i - 1, 2].PutValue(DateTime.Today.AddDays(i)); // Date
            }

            // ----- Duplicate the header row to a new position (row 1) -----
            // CopyRow(sourceCells, sourceRowIndex, destinationRowIndex)
            cells.CopyRow(cells, 0, 1);

            // ----- Freeze the two header rows so they stay visible while scrolling -----
            // FreezePanes(row, column, freezedRows, freezedColumns)
            // Row index 2 means the freeze line is just above row 2, thus rows 0 and 1 are frozen.
            sheet.FreezePanes(2, 0, 2, 0);

            // Save the workbook
            workbook.Save("HeaderCopyAndFreeze.xlsx");
        }
    }
}
