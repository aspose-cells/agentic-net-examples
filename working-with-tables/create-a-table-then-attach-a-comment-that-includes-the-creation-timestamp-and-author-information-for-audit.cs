// Title: Add author and timestamp audit comment to an Excel ListObject with Aspose.Cells for .NET (C#)
// Description: This C# example demonstrates how to create a new workbook, insert a ListObject (Excel table) over a data range, assign a table‑level comment that records the author name and creation date‑time, and add a threaded comment to a specific cell inside the table for richer audit metadata before saving the file as .xlsx.
// Keywords: Aspose.Cells C# ListObject comment | Excel table audit comment Aspose | add threaded comment Aspose.Cells | author timestamp comment Excel | Aspose.Cells table metadata | C# create Excel table with comment | Aspose.Cells audit trail
// Common Searches: how to add comment to ListObject Aspose.Cells | Aspose.Cells add author timestamp to Excel table | C# threaded comment inside Excel table Aspose | set audit information on Excel table using Aspose | Aspose.Cells example for table comments
// Developer Intent: Create a ListObject and embed audit information (author, creation time) via table comment and cell threaded comment.
// Use Cases: Compliance reports that require embedded audit trails for each data table. | Automated generation of workbooks where version and author metadata are stored directly in the Excel file. | Change‑log tracking by attaching threaded comments to key cells within a table. | Regulatory documentation that mandates author and timestamp on data tables.
// AI Prompts: Generate C# code with Aspose.Cells to create an Excel table and set its Comment property to include a given author and current datetime. | Show how to add a threaded comment to a cell inside a ListObject, specifying author email and role, using Aspose.Cells for .NET. | Explain how to read the table comment and threaded comment after opening the workbook with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableCommentDemo
{
    // This C# example demonstrates how to create a new workbook, insert a ListObject (Excel table) over a data range, assign a table‑level comment that records the author name and creation date‑time, and add a threaded comment to a specific cell inside the table for richer audit metadata before saving the file as .xlsx.
    public class Program
    {
        public static void Main()
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

            // Add a ListObject (table) covering the data range
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
            int tableIndex = worksheet.ListObjects.Add(0, 0, 2, 1, true);
            ListObject table = worksheet.ListObjects[tableIndex];
            table.DisplayName = "EmployeeTable";

            // Define audit information
            string author = "John Doe";
            DateTime createdTime = DateTime.Now;

            // Attach an audit comment to the table
            table.Comment = $"Created by {author} on {createdTime:yyyy-MM-dd HH:mm:ss}";

            // Optionally, you can also add a threaded comment to a cell inside the table for richer metadata
            // Add a threaded comment author
            int authorId = workbook.Worksheets.ThreadedCommentAuthors.Add(author, "john.doe@example.com", "PROVIDER");
            ThreadedCommentAuthor threadedAuthor = workbook.Worksheets.ThreadedCommentAuthors[authorId];

            // Add a threaded comment to the first data cell (A2)
            int commentIdx = worksheet.Comments.Add(1, 0); // Row 1, Column 0 => A2
            Comment comment = worksheet.Comments[commentIdx];
            int threadedIdx = comment.ThreadedComments.Add("Table created for audit purposes.", threadedAuthor);
            ThreadedComment threadedComment = comment.ThreadedComments[threadedIdx];
            threadedComment.CreatedTime = createdTime;

            // Save the workbook
            workbook.Save("TableWithAuditComment.xlsx", SaveFormat.Xlsx);
        }
    }
}
