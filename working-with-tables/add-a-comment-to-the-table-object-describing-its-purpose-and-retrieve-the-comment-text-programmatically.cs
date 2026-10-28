// Title: Add and read a comment on an Aspose.Cells ListObject (Excel table) using C#
// Common Searches: Aspose.Cells C# add comment to ListObject table | retrieve comment from Excel table with Aspose.Cells .NET | set ListObject.Comment property programmatically in C# | read table comment after saving workbook using Aspose.Cells | how to use ListObject.Comment with Aspose.Cells for .NET
// Tags: Aspose.Cells ListObject comment property | C# add comment to Excel ListObject | retrieve ListObject comment .NET | Excel table metadata with Aspose.Cells | save workbook with table comment C#

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The sample creates a workbook, defines a data range, adds a ListObject named "SalesTable", assigns a comment describing its purpose, retrieves the comment via the ListObject.Comment property, prints it, and saves the file as TableWithComment.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Fill sample data that will become the table
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(20);

            // Define the range that will be converted into a table (including header)
            int firstRow = 0;      // zero‑based index for row 1
            int firstColumn = 0;   // zero‑based index for column A
            int totalRows = 3;     // rows 1‑3
            int totalColumns = 2;  // columns A‑B

            // Add a ListObject (table) to the worksheet; the last argument indicates that the range has headers
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Set the table name (DisplayName is the correct property)
            table.DisplayName = "SalesTable";

            // Add a comment describing the purpose of the table
            table.Comment = "This table stores product sales quantities.";

            // Retrieve the comment programmatically
            string retrievedComment = table.Comment;
            Console.WriteLine("Table Comment: " + retrievedComment);

            // Determine output file path and ensure the directory exists
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithComment.xlsx");
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved to: " + outputPath);
            }
            catch (Exception saveEx)
            {
                Console.WriteLine("Failed to save workbook: " + saveEx.Message);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
