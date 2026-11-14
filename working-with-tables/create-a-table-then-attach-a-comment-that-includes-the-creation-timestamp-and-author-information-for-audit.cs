// Title: Create an Excel ListObject table with a header row and add an audit comment (author and timestamp) using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add ListObject table with header row | how to attach author and timestamp comment to a cell using Aspose.Cells | C# create Excel table and add audit comment with date in Aspose.Cells | Aspose.Cells insert comment with dynamic timestamp on worksheet cell | example of saving workbook with table and comment using Aspose.Cells .NET
// Tags: add ListObject table Aspose.Cells C# | insert cell comment with author timestamp Aspose.Cells | create Excel table with header row Aspose.Cells | audit comment on worksheet cell Aspose.Cells | save workbook as .xlsx Aspose.Cells C#

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;

// The sample demonstrates how to create a new workbook, define a range as a ListObject table named "SampleTable" with a header row, add an audit comment containing the author name and current timestamp to cell A1, and save the file as "TableWithComment.xlsx" using Aspose.Cells for .NET.
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

            // Populate data for the table
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");

            // Define the range for the table (A1:B3)
            int firstRow = 0;      // zero‑based index for row 1
            int firstColumn = 0;   // zero‑based index for column A
            int totalRows = 3;     // includes header row
            int totalColumns = 2;  // columns A and B

            // Add a ListObject (table) to the worksheet; the last parameter indicates that the range has a header row
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "SampleTable";
            table.ShowHeaderRow = true;
            // Note: ShowTotalRow property is not available in this version of Aspose.Cells; omitted.

            // Prepare audit information
            string author = "John Doe";
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            string commentText = $"Created by {author} on {timestamp}";

            // Add a comment to cell A1 with the audit details
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = commentText;
            comment.Author = author;

            // Save the workbook to a file
            workbook.Save("TableWithComment.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
