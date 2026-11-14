// Title: Create an Excel ListObject table and add a formatted cell comment with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# example to create a ListObject table and insert a comment | How to set comment font color and size in an Excel file using Aspose.Cells for .NET | Add a blue colored note to cell A1 when generating an Excel workbook with Aspose.Cells | C# code to generate an Employees table and attach a formatted comment in Excel | Save a workbook with a table and custom comment using Aspose.Cells .NET API
// Tags: Aspose.Cells create ListObject table | Aspose.Cells add cell comment | Aspose.Cells format comment font | Aspose.Cells save workbook as xlsx | C# generate Excel table with comment

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Drawing;

// The sample program creates a new workbook, populates cells A1:B3 with sample data, defines a ListObject named "Employees" covering that range, adds a comment to cell A1 with the note "Reference Document", formats the comment text to blue color and size 10, and saves the file as TableWithComment.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet ws = workbook.Worksheets[0];

            // Populate some data that will become a table
            ws.Cells["A1"].PutValue("ID");
            ws.Cells["B1"].PutValue("Name");
            ws.Cells["A2"].PutValue(1);
            ws.Cells["B2"].PutValue("Alice");
            ws.Cells["A3"].PutValue(2);
            ws.Cells["B3"].PutValue("Bob");

            // Create a table (ListObject) covering the data range A1:B3
            int firstRow = 0;          // zero‑based index for row 1
            int firstColumn = 0;       // zero‑based index for column A
            int totalRows = 3;         // rows A1:B3 inclusive
            int totalColumns = 2;
            int tableIndex = ws.ListObjects.Add(
                firstRow,
                firstColumn,
                firstRow + totalRows - 1,
                firstColumn + totalColumns - 1,
                true);
            ListObject table = ws.ListObjects[tableIndex];
            table.DisplayName = "Employees";
            table.ShowHeaderRow = true;

            // Add a comment to cell A1
            int commentIndex = ws.Comments.Add("A1");
            Comment comment = ws.Comments[commentIndex];

            // Set comment text (use Note property)
            comment.Note = "Reference Document";

            // Optional formatting for the comment text
            comment.Font.Color = Color.Blue;
            comment.Font.Size = 10;

            // Save the workbook to a file
            string outputPath = "TableWithComment.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
