// Title: Copy rows in Aspose.Cells for .NET while preserving hidden row state and row height
// AI Prompts: Write C# code using Aspose.Cells to duplicate a block of rows to a new location and copy each row's IsHidden property and Height. | Create a function that copies rows 0‑4 to row 6 in a worksheet, ensuring hidden flags and row heights are retained with Aspose.Cells.
// Common Searches: how to copy rows and keep hidden rows in Aspose.Cells C# | Aspose.Cells copy rows preserve row visibility and height | duplicate rows with hidden state using Aspose.Cells for .NET | copy multiple rows and retain hidden flag Aspose.Cells example
// Tags: Aspose.Cells copy rows with hidden state | C# duplicate rows preserving row height Aspose.Cells | preserve row visibility during copy Aspose.Cells | copy rows and maintain formatting Aspose.Cells .NET | Aspose.Cells copy rows example with hidden rows

using System;
using Aspose.Cells;

// The sample creates a workbook, hides rows 2 and 4, copies rows 0‑4 to start at row 6, then transfers each source row's IsHidden flag and Height to the corresponding target row, and finally saves the file as CopyRowsPreserveHidden.xlsx.
class CopyRowsPreserveHidden
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data in rows 0‑4
        for (int i = 0; i < 5; i++)
        {
            cells[i, 0].PutValue($"Row {i + 1}");
        }

        // Hide specific rows (row indices are zero‑based)
        worksheet.Cells.HideRow(1); // hide row 2
        worksheet.Cells.HideRow(3); // hide row 4

        // Parameters for copying
        int sourceRowIndex = 0;      // first row to copy
        int destinationRowIndex = 6; // where to paste the copied rows
        int rowNumber = 5;           // number of rows to copy

        // Copy rows data and formats
        worksheet.Cells.CopyRows(worksheet.Cells, sourceRowIndex, destinationRowIndex, rowNumber);

        // Preserve hidden state (and optionally height) for each copied row
        for (int i = 0; i < rowNumber; i++)
        {
            Row sourceRow = worksheet.Cells.Rows[sourceRowIndex + i];
            Row targetRow = worksheet.Cells.Rows[destinationRowIndex + i];

            // Copy hidden flag
            targetRow.IsHidden = sourceRow.IsHidden;

            // Copy row height (optional, keeps visual appearance identical)
            targetRow.Height = sourceRow.Height;
        }

        // Save the workbook
        workbook.Save("CopyRowsPreserveHidden.xlsx");
    }
}
