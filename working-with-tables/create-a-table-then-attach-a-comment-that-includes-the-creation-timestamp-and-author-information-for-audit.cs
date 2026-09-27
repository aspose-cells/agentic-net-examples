// Title: Generate an Excel ListObject table and embed an audit comment with author and timestamp using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a worksheet, adds a ListObject table over a data range, and inserts a comment on the header cell containing a given author name and the current date‑time. | Extend the Aspose.Cells example to include a total row in the ListObject and format the audit comment so the author and timestamp appear in bold. | Create a reusable C# method that accepts author, data range, and output path, then builds a ListObject table, attaches the audit comment to the first header cell, and saves the workbook.
// Common Searches: how to add a ListObject table with a header comment in Aspose.Cells C# | Aspose.Cells create Excel table and set author timestamp comment on cell A1 | C# Aspose.Cells add audit metadata comment to worksheet header row
// Tags: Aspose.Cells create ListObject table | Aspose.Cells add cell comment with timestamp | C# generate Excel table with audit metadata | Aspose.Cells save workbook as .xlsx with audit comment | Aspose.Cells ListObject total row customization

using System;
using Aspose.Cells;
using Aspose.Cells.Tables; // Required for ListObject

// Demonstrates using Aspose.Cells for .NET to create a new workbook, add a ListObject (Excel table) with headers and sample data, attach a comment containing the author name and current timestamp to the first header cell, and save the file as AuditTable.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate header row
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");

            // Populate sample data rows
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["C2"].PutValue(85);

            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["C3"].PutValue(92);

            // Define the range for the table (including header)
            int firstRow = 0;          // zero‑based index for row 1
            int firstColumn = 0;       // zero‑based index for column A
            int totalRows = 3;         // header + 2 data rows
            int totalColumns = 3;      // ID, Name, Score

            // Add a ListObject (Excel table) to the worksheet
            int tableIndex = sheet.ListObjects.Add(
                firstRow,
                firstColumn,
                firstRow + totalRows - 1,
                firstColumn + totalColumns - 1,
                true);

            ListObject table = sheet.ListObjects[tableIndex];
            // Use DisplayName instead of Name (compatible with all versions)
            table.DisplayName = "AuditTable";
            table.ShowHeaderRow = true;
            // ShowTotalRow property may not be available in older versions; omit if not present

            // Prepare audit comment with timestamp and author
            string author = "John Doe";
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            string commentText = $"Created by {author} on {timestamp}";

            // Attach the comment to the first header cell (A1)
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = commentText;
            comment.Author = author;

            // Save the workbook to a file
            workbook.Save("AuditTable.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
