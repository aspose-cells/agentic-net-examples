// Title: Copy Rows with Hyperlinks and Preserve Targets Using Aspose.Cells for .NET
// Description: Demonstrates how to copy a worksheet row that contains a hyperlink with Aspose.Cells. The example shows the default copy behavior that creates a new hyperlink entry, then uses CopyOptions.ExtendToAdjacentRange to extend the original hyperlink range instead of duplicating it. It prints the hyperlink count, address, and start/end rows, and saves the workbook for verification.
// Keywords: Aspose.Cells copy rows | hyperlink copy options | ExtendToAdjacentRange | C# Aspose.Cells example | preserve hyperlink after row copy | worksheet hyperlink range | CopyRows method | .NET spreadsheet library
// Common Searches: Aspose.Cells copy row with hyperlink | ExtendToAdjacentRange hyperlink behavior | how to keep hyperlink target when copying rows | hyperlink count after copying rows Aspose.Cells | C# copy rows without duplicating hyperlinks
// Developer Intent: Copy rows that contain hyperlinks and ensure the hyperlinks continue to point to the original target after the copy operation.
// Use Cases: Duplicate a data row with its hyperlink, creating a separate hyperlink entry (default copy). | Copy a row while extending the existing hyperlink range so the original hyperlink covers both rows. | Programmatically verify hyperlink address and range after copying rows to confirm correct behavior.
// AI Prompts: Write C# code using Aspose.Cells that copies a row with a hyperlink and uses CopyOptions.ExtendToAdjacentRange to extend the hyperlink range instead of creating a new entry. | Create a method that copies rows, returns the hyperlink count before and after the copy, and checks that the hyperlink address remains unchanged. | Explain the effect of the ExtendToAdjacentRange property on hyperlink handling during row copy operations in Aspose.Cells, and show sample console output of hyperlink area properties.

using System;
using Aspose.Cells;

// Demonstrates how to copy a worksheet row that contains a hyperlink with Aspose.Cells. The example shows the default copy behavior that creates a new hyperlink entry, then uses CopyOptions.ExtendToAdjacentRange to extend the original hyperlink range instead of duplicating it. It prints the hyperlink count, address, and start/end rows, and saves the workbook for verification.
class CopyRowsWithHyperlinksDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // ------------------------------------------------------------
        // 1. Add a hyperlink in row 1 (zero‑based index) – cell A2
        // ------------------------------------------------------------
        ws.Cells["A2"].PutValue("Row with hyperlink");
        ws.Hyperlinks.Add("A2", 1, 1, "https://www.example.com");

        // ------------------------------------------------------------
        // 2. Copy the row without any CopyOptions (default behavior)
        //    This creates a new hyperlink for the copied row.
        // ------------------------------------------------------------
        ws.Cells.CopyRows(ws.Cells, 1, 2, 1);
        Console.WriteLine("Hyperlink count after default copy: " + ws.Hyperlinks.Count);
        // Expected: 2 (original + copied)

        // ------------------------------------------------------------
        // 3. Reset worksheet to test ExtendToAdjacentRange option
        // ------------------------------------------------------------
        ws.Cells.ClearContents(0, 0, ws.Cells.MaxRow + 1, ws.Cells.MaxColumn + 1);
        ws.Hyperlinks.Clear();

        ws.Cells["A2"].PutValue("Row with hyperlink");
        ws.Hyperlinks.Add("A2", 1, 1, "https://www.example.com");

        // ------------------------------------------------------------
        // 4. Copy the row using CopyOptions.ExtendToAdjacentRange = true
        //    The hyperlink range is extended; no new hyperlink is added.
        // ------------------------------------------------------------
        CopyOptions options = new CopyOptions();
        options.ExtendToAdjacentRange = true;

        ws.Cells.CopyRows(ws.Cells, 1, 2, 1, options);

        // ------------------------------------------------------------
        // 5. Verify results
        // ------------------------------------------------------------
        Console.WriteLine("Hyperlink count after copy with ExtendToAdjacentRange: " + ws.Hyperlinks.Count);
        // Expected: 1 (range extended, not duplicated)

        Hyperlink link = ws.Hyperlinks[0];

        // Address should remain unchanged
        Console.WriteLine("Hyperlink address: " + link.Address);

        // The area should now span rows 1 and 2 (zero‑based indices)
        Console.WriteLine("Hyperlink start row (zero‑based): " + link.Area.StartRow);
        Console.WriteLine("Hyperlink end row (zero‑based): " + link.Area.EndRow);
        // Expected EndRow = 2

        // ------------------------------------------------------------
        // 6. Save the workbook for visual inspection if needed
        // ------------------------------------------------------------
        wb.Save("CopyRowsHyperlinkDemo.xlsx");
    }
}
