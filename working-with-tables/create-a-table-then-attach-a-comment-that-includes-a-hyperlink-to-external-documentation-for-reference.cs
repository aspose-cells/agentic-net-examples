// Title: Create an Excel ListObject table with a header comment that links to Aspose.Cells documentation using C#
// AI Prompts: Generate C# code that adds a ListObject to the first worksheet, inserts a comment in cell A1, and sets a hyperlink on the comment shape pointing to the Aspose.Cells online docs. | Show how to configure the comment hyperlink's screen tip and enable row‑stripe styling for the table while saving the workbook as TableWithComment.xlsx. | Provide a C# example that creates sample data, defines a table range, adds a comment with a hyperlink, and applies custom table style options in Aspose.Cells.
// Common Searches: Aspose.Cells C# add comment with hyperlink to documentation in table header | how to set hyperlink on comment shape in Aspose.Cells .NET | C# create ListObject and attach external link to comment in Excel file | apply row stripe style to Aspose.Cells table while preserving comment hyperlink
// Tags: listobject table with comment hyperlink Aspose.Cells | set comment shape hyperlink C# | apply row stripe style Aspose.Cells .NET | add header comment linking to external docs Excel | save workbook with table and comment hyperlink Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;   // Required for ListObject

namespace AsposeCellsExample
{
    // Demonstrates creating a workbook, adding a ListObject table with sample data, inserting a comment in the header cell, attaching a hyperlink to the comment shape that points to the Aspose.Cells documentation, applying row‑stripe styling, and saving the file as TableWithComment.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("ID");
                sheet.Cells["B1"].PutValue("Name");
                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue("Alice");
                sheet.Cells["A3"].PutValue(2);
                sheet.Cells["B3"].PutValue("Bob");

                // Define the range for the table (including header)
                int firstRow = 0;          // zero‑based index for row 1
                int firstColumn = 0;       // zero‑based index for column A
                int totalRows = 3;         // rows count (including header)
                int totalColumns = 2;      // columns count

                // Add a ListObject (table) to the worksheet
                int tableIndex = sheet.ListObjects.Add(
                    firstRow,
                    firstColumn,
                    firstRow + totalRows - 1,
                    firstColumn + totalColumns - 1,
                    true);

                ListObject table = sheet.ListObjects[tableIndex];
                table.DisplayName = "SampleTable";
                table.ShowTableStyleFirstColumn = false;
                table.ShowTableStyleLastColumn = false;
                table.ShowTableStyleRowStripes = true;
                table.ShowTableStyleColumnStripes = false;

                // Add a comment to the top‑left cell of the table (A1)
                int commentIndex = sheet.Comments.Add("A1");
                Comment comment = sheet.Comments[commentIndex];
                comment.Note = "See external documentation for details.";

                // Attach a hyperlink to the comment shape
                // The Hyperlink property is read‑only; modify its fields instead of assigning a new object
                Hyperlink hl = comment.CommentShape.Hyperlink;
                hl.Address = "https://docs.aspose.com/cells/net/";
                hl.ScreenTip = "Aspose.Cells Documentation";

                // Save the workbook
                workbook.Save("TableWithComment.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
