// Title: Generate an Excel ListObject table and add a timestamped author comment with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to create a ListObject table, fill it with sample rows, and insert a comment on cell A1 that includes the current date, time, and Windows user name. | Show how to add an audit‑trail comment to the header cell of an Excel table in a workbook created with Aspose.Cells for .NET. | Demonstrate saving the workbook as an .xlsx file after adding a ListObject and a timestamped comment in C#. | Explain how to retrieve the current environment username and embed it in a cell comment with Aspose.Cells.
// Common Searches: aspocells c# create listobject table with header and add comment containing date and user | how to programmatically add an audit comment to Excel cell using Aspose.Cells .NET | C# Aspose.Cells add timestamped comment to cell A1 in new workbook | save workbook as AuditTable.xlsx after inserting ListObject and comment Aspose.Cells
// Tags: Aspose.Cells ListObject creation C# | Aspose.Cells add cell comment with timestamp | Excel table header row ListObject Aspose | audit comment author environment username Aspose.Cells | save workbook as .xlsx Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Creates a new workbook, defines a two‑column ListObject named 'AuditTable' with sample data, adds a comment to cell A1 that records the current date, time and Windows user, and saves the file as AuditTable.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the table
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("John");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Jane");

            // Define the range of the table (including header row)
            int firstRow = 0;      // zero‑based index for row 1
            int firstColumn = 0;   // zero‑based index for column A
            int totalRows = 3;     // rows 1‑3
            int totalColumns = 2;  // columns A‑B

            // Add a ListObject (Excel table) to the worksheet
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "AuditTable";
            table.ShowHeaderRow = true;
            // ShowTotalRow property is not available in this version; default is false.

            // Create a comment on the top‑left cell (A1) with timestamp and author
            string cellName = "A1";
            int commentIndex = sheet.Comments.Add(cellName);
            Comment comment = sheet.Comments[commentIndex];
            comment.Author = Environment.UserName;
            comment.Note = $"Created on {DateTime.Now:yyyy-MM-dd HH:mm:ss} by {Environment.UserName}";

            // Save the workbook to a file
            string outputPath = "AuditTable.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
