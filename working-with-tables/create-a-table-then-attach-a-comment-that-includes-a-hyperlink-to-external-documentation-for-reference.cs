// Title: Create an Excel ListObject table and add a hyperlink comment to a cell with Aspose.Cells for .NET (C#)
// Common Searches: how to create an Excel ListObject table with Aspose.Cells in C# | adding a comment that contains a URL to a specific cell using Aspose.Cells .NET | example of setting the Note property for a comment in Aspose.Cells C# | saving a workbook after inserting a table and a hyperlink comment with Aspose.Cells | checking and creating the output folder before saving an Excel file in Aspose.Cells C#
// Tags: Aspose.Cells add ListObject to worksheet C# | Aspose.Cells set comment Note property C# | Aspose.Cells embed hyperlink in cell comment | Aspose.Cells save workbook with table and comment | Aspose.Cells ensure output directory exists C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example demonstrates how to: | Step | Action | |------|--------| | 1 | Instantiate a `Workbook` and obtain the first `Worksheet`. | | 2 | Populate header and data cells (ID, Name). | | 3 | Define the range and add a `ListObject` (table) with headers. | | 4 | Add a comment to cell **B2** and set its `Note` property to a documentation URL. | | 5 | Verify the output directory exists and save the workbook as `TableWithComment.xlsx`. | Reference: [Aspose.Cells .NET documentation](https://docs.aspose.com/cells/net/)
class TableWithCommentExample
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
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");

            // Define the range of the table (including header row)
            int firstRow = 0;      // zero‑based index
            int firstColumn = 0;
            int totalRows = 3;     // header + 2 data rows
            int totalColumns = 2;

            // Add a ListObject (table) to the worksheet; the last argument indicates that the range has headers
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "SampleTable";   // set a friendly name
            table.ShowHeaderRow = true;
            table.ShowTotals = false;            // hide the total row

            // Add a simple comment to cell B2 (row 1, column 1)
            int commentRow = 1;
            int commentColumn = 1;
            int commentIndex = sheet.Comments.Add(commentRow, commentColumn);
            Comment comment = sheet.Comments[commentIndex];
            // Use the Note property to set the comment text (compatible with current Aspose.Cells versions)
            comment.Note = "For more information, see the documentation: https://example.com/docs";

            // Save the workbook to a file
            string outputPath = "TableWithComment.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
