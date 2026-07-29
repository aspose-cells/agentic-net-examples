// Title: Copy a Header Row and Freeze Top Rows with Cells.CopyRow in Aspose.Cells for .NET
// Description: Demonstrates how to use Aspose.Cells for .NET to copy the first worksheet row, create a second header, and freeze the top two rows so both headers stay visible while scrolling.
// Keywords: Aspose.Cells | C# copy row | Cells.CopyRow | FreezePanes | duplicate header | freeze top rows | Excel header repeat | Aspose.Cells .NET example | copy row with formatting | freeze panes C#
// Common Searches: Aspose.Cells copy first row to second row C# | How to freeze top rows in Aspose.Cells | Duplicate header and freeze panes using Aspose.Cells | Cells.CopyRow example .NET | FreezePanes after copying row Aspose.Cells
// Developer Intent: Create a second header row by copying row 0 and keep both rows fixed with FreezePanes.
// Use Cases: Generating reports where the header must appear twice for printing on each page. | Building large data tables where users need a persistent multi‑row header while scrolling. | Applying identical formatting to a secondary header without redefining styles manually.
// AI Prompts: Provide a C# Aspose.Cells snippet that copies row 0 to row 1, preserves formatting, and freezes the first two rows. | Show how to use Cells.CopyRow and Worksheet.FreezePanes together to keep duplicated headers visible during scroll. | Explain step‑by‑step how to duplicate a header row and set FreezePanes for the top two rows in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Demonstrates how to use Aspose.Cells for .NET to copy the first worksheet row, create a second header, and freeze the top two rows so both headers stay visible while scrolling.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // ----- Original header row (row 0) -----
        cells["A1"].PutValue("ID");
        cells["B1"].PutValue("Name");
        cells["C1"].PutValue("Date");

        // ----- Sample data rows (starting from row 2) -----
        for (int i = 2; i <= 10; i++)
        {
            int rowIndex = i - 1; // zero‑based index
            cells[rowIndex, 0].PutValue(i - 1);                     // ID
            cells[rowIndex, 1].PutValue("Item " + (i - 1));        // Name
            cells[rowIndex, 2].PutValue(DateTime.Now.AddDays(i-2)); // Date
        }

        // ----- Duplicate the header row -----
        // Copy the whole first row (index 0) to the second row (index 1)
        cells.CopyRow(cells, 0, 1);

        // ----- Freeze panes so that both header rows stay visible while scrolling -----
        // Freeze the top two rows (row index 2 is the first scrollable row)
        worksheet.FreezePanes(2, 0, 2, 0);

        // Save the workbook
        workbook.Save("HeaderCopyAndFreeze.xlsx");
    }
}
