// Title: Add an audit comment with author and timestamp to an Excel table using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, insert sample data, define a ListObject (Excel table), set its display name, generate an audit string containing the creator's name and the current date‑time, assign it to the table's Comment property, and save the file as XLSX.
// Keywords: Aspose.Cells C# add table comment | Excel ListObject audit metadata .NET | table comment with author timestamp | programmatically set Excel table comment | Aspose.Cells create table with audit trail
// Common Searches: Aspose.Cells add comment to ListObject | C# set author and creation date on Excel table | How to attach audit information to an Excel table using Aspose | Save Excel table with metadata in .NET
// Developer Intent: Create a ListObject in a worksheet and attach a comment that records who created the table and when.
// Use Cases: Compliance reporting where each table must show its creator and creation time. | Automated spreadsheet generation that embeds ownership metadata for version control. | Shared workbook templates that log author information for multi‑user auditing.
// AI Prompts: Generate C# code with Aspose.Cells that adds a table and writes an audit comment containing the current user and timestamp. | Show how to retrieve, modify, or delete the comment of an existing Aspose.Cells ListObject. | Explain best practices for storing audit metadata in Excel tables using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Demonstrates how to create a workbook, insert sample data, define a ListObject (Excel table), set its display name, generate an audit string containing the creator's name and the current date‑time, assign it to the table's Comment property, and save the file as XLSX.
class TableWithAuditComment
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the table (including header row)
        worksheet.Cells["A1"].PutValue("ID");
        worksheet.Cells["B1"].PutValue("Name");
        worksheet.Cells["A2"].PutValue(1);
        worksheet.Cells["B2"].PutValue("Alice");
        worksheet.Cells["A3"].PutValue(2);
        worksheet.Cells["B3"].PutValue("Bob");

        // Add a ListObject (table) that covers the data range A1:B3
        // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
        int tableIndex = worksheet.ListObjects.Add(0, 0, 2, 1, true);
        ListObject table = worksheet.ListObjects[tableIndex];
        table.DisplayName = "Employees";

        // Prepare audit information
        string author = "John Doe";
        string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

        // Attach a comment to the table containing author and creation timestamp
        table.Comment = $"Created by {author} on {timestamp}";

        // Save the workbook to an XLSX file
        workbook.Save("TableWithAuditComment.xlsx", SaveFormat.Xlsx);
    }
}
