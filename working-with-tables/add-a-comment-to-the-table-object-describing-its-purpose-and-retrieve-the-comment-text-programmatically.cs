// Title: Add and retrieve a comment on an Aspose.Cells ListObject (Excel table) using C#
// Common Searches: how to set a comment on an Excel table (ListObject) with Aspose.Cells in C# | Aspose.Cells C# retrieve ListObject comment programmatically | example of adding metadata comment to a table using Aspose.Cells .NET | read table comment from a workbook created with Aspose.Cells | Aspose.Cells ListObject comment property usage in C# code
// Tags: Aspose.Cells ListObject add comment C# | Aspose.Cells retrieve table comment .NET | Excel table comment handling with Aspose.Cells API | set ListObject comment programmatically | read ListObject comment property | Aspose.Cells workbook save with table comment

using System;
using Aspose.Cells;
using Aspose.Cells.Tables; // Required for ListObject

// The example creates a workbook, adds a ListObject named SalesTable, assigns a descriptive comment via the Comment property, retrieves the comment back into a string, outputs it to the console, and saves the workbook as TableWithComment.xlsx.
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

            // Populate data that will be part of the table
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(50);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(30);

            // Define the range for the table (including header row)
            int firstRow = 0;          // zero‑based index for row 1
            int firstColumn = 0;       // zero‑based index for column A
            int totalRows = 3;         // header + 2 data rows
            int totalColumns = 2;      // two columns: Product and Quantity

            // Add a ListObject (table) to the worksheet
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn, firstRow + totalRows, firstColumn + totalColumns, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "SalesTable";

            // Add a comment describing the purpose of the table
            table.Comment = "This table stores product sales quantities.";

            // Retrieve the comment programmatically
            string retrievedComment = table.Comment;

            // Output the comment to the console
            Console.WriteLine("Table Comment: " + retrievedComment);

            // Save the workbook (optional)
            workbook.Save("TableWithComment.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
