// Title: Duplicate Header Row and Freeze It with Aspose.Cells for .NET (C#)
// Description: Learn how to copy a header row to another position using Cells.CopyRow and then lock the copied row with Worksheet.FreezePanes so it stays visible while scrolling in an Excel workbook created with Aspose.Cells for .NET.
// Keywords: Aspose.Cells CopyRow C# | Aspose.Cells FreezePanes | duplicate header row Excel | copy row and freeze pane .NET | Aspose.Cells example C# | Excel header repeat Aspose | freeze panes after copy row
// Common Searches: Aspose.Cells copy header row and freeze | C# Cells.CopyRow example | How to freeze a copied row in Aspose.Cells | Duplicate and lock header in Excel using Aspose | CopyRow then FreezePanes Aspose.Cells
// Developer Intent: Copy an existing header row to a new location and freeze that row so it remains visible during scrolling.
// Use Cases: Create a printable report with a repeated header at the bottom of a long data set while keeping the duplicated header fixed on screen. | Build a dashboard worksheet where the header appears both at the top and near the data summary, staying visible as users scroll. | Generate multi‑page Excel files where each page starts with the same styled header that is also frozen for on‑screen navigation.
// AI Prompts: Provide C# code that uses Aspose.Cells Cells.CopyRow to copy a header row and then applies Worksheet.FreezePanes to keep the copied header visible. | Show an Aspose.Cells for .NET example that duplicates a header row at a specific index and freezes the pane at that row.

using System;
using Aspose.Cells;

// Learn how to copy a header row to another position using Cells.CopyRow and then lock the copied row with Worksheet.FreezePanes so it stays visible while scrolling in an Excel workbook created with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate the original header row (row 0) and some sample data rows
        cells["A1"].PutValue("ID");
        cells["B1"].PutValue("Name");
        cells["C1"].PutValue("Score");
        for (int i = 1; i <= 10; i++)
        {
            cells[i, 0].PutValue(i);                     // ID
            cells[i, 1].PutValue("Student " + i);        // Name
            cells[i, 2].PutValue(50 + i);                // Score
        }

        // Define source (original header) and destination row indices (zero‑based)
        int sourceRowIndex = 0;   // Row 0 contains the header
        int destinationRowIndex = 12; // Copy header to row 13 (index 12)

        // Duplicate the header row using Cells.CopyRow
        cells.CopyRow(cells, sourceRowIndex, destinationRowIndex);

        // Freeze panes so that the copied header stays visible while scrolling
        // Freeze all rows up to and including the copied header (destinationRowIndex + 1 rows)
        worksheet.FreezePanes(destinationRowIndex + 1, 0, destinationRowIndex + 1, 0);

        // Save the workbook to a file
        workbook.Save("HeaderCopyAndFreeze.xlsx");
    }
}
