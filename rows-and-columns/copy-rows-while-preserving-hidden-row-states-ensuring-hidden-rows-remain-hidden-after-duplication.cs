// Title: Aspose.Cells for .NET – Copy rows while keeping hidden rows hidden (C#)
// Description: Demonstrates how to copy a block of rows to another location using Cells.CopyRows and then transfer the IsHidden flag so that any rows originally hidden stay hidden in the duplicated range. The example creates a workbook, fills rows, hides specific rows, copies them, syncs visibility, and saves the file.
// Keywords: Aspose.Cells copy rows C# | preserve hidden rows Aspose.Cells | CopyRows IsHidden property | duplicate rows keep hidden state | Excel hidden rows automation .NET | Aspose.Cells row visibility | C# Excel copy rows hidden | Aspose.Cells hidden row handling | Excel worksheet row copy preserve format
// Common Searches: Aspose.Cells copy rows keep hidden rows hidden | How to preserve hidden row state after copying rows in Aspose.Cells | Copy rows with hidden rows using Aspose.Cells .NET | Aspose.Cells CopyRows retain IsHidden flag | C# copy Excel rows and keep hidden rows hidden
// Developer Intent: Copy a range of rows to a new location and ensure that any rows that were hidden in the source range remain hidden in the destination range.
// Use Cases: Replicate a template section that includes collapsed group headers, preserving the hidden state for each copy. | Generate a printable report by moving data blocks while maintaining hidden rows used for subtotal grouping. | Programmatically duplicate filtered data where hidden rows represent collapsed details, keeping the same visibility after copy.
// AI Prompts: Write C# code with Aspose.Cells to copy rows 10‑20 to rows 30‑40 and retain hidden rows. | Show an Aspose.Cells .NET example that copies rows and synchronizes the IsHidden property after using CopyRows. | Explain step‑by‑step how to copy rows with formatting and preserve hidden row state in Aspose.Cells without affecting other worksheet settings.

using System;
using Aspose.Cells;

namespace AsposeCellsRowCopyPreserveHidden
{
    // Demonstrates how to copy a block of rows to another location using Cells.CopyRows and then transfer the IsHidden flag so that any rows originally hidden stay hidden in the duplicated range. The example creates a workbook, fills rows, hides specific rows, copies them, syncs visibility, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data in rows 0-5
            for (int r = 0; r < 6; r++)
            {
                cells[r, 0].PutValue($"Row {r + 1} Col A");
                cells[r, 1].PutValue($"Row {r + 1} Col B");
            }

            // Hide rows 2 and 4 (zero‑based indexes 1 and 3)
            sheet.Cells.HideRow(1);
            sheet.Cells.HideRow(3);

            // Destination start row (copy rows 0‑5 to rows 6‑11)
            int sourceStart = 0;
            int destinationStart = 6;
            int rowCount = 6;

            // Copy rows data and formats
            sheet.Cells.CopyRows(sheet.Cells, sourceStart, destinationStart, rowCount);

            // Preserve hidden state for each copied row
            for (int i = 0; i < rowCount; i++)
            {
                Row sourceRow = sheet.Cells.Rows[sourceStart + i];
                Row destRow = sheet.Cells.Rows[destinationStart + i];
                destRow.IsHidden = sourceRow.IsHidden;
            }

            // Save the workbook
            workbook.Save("RowsCopiedPreserveHidden.xlsx");
        }
    }
}
