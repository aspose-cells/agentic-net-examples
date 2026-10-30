// Title: Create an Excel ListObject table and add a visible comment containing a hyperlink using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# how to add a ListObject table and display a comment containing a URL | programmatically create an Excel table and attach a visible comment with a link using Aspose.Cells for .NET | C# example for saving a workbook that includes a table and a cell comment with a hyperlink | add a comment with a hyperlink to cell A1 in an Aspose.Cells worksheet | make a cell comment visible by default when generating Excel files with Aspose.Cells
// Tags: create ListObject table Aspose.Cells | insert cell comment with URL C# | set comment visibility Aspose.Cells | populate worksheet range for table Aspose.Cells | save Excel file with table and comment Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The sample program creates a new workbook, fills cells A1:B3 with sample data, defines that range as a ListObject table named "SampleTable", adds a visible comment to cell A1 that includes a documentation hyperlink, ensures the output directory exists, saves the file as TableWithComment.xlsx, and writes a success message to the console.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data that will be converted into a table
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");

            // Define the range for the table (A1:B3)
            int firstRow = 0;          // zero‑based index for row 1
            int firstColumn = 0;       // zero‑based index for column A
            int totalRows = 3;         // rows 1‑3
            int totalColumns = 2;      // columns A‑B

            // Add a ListObject (table) to the worksheet
            int tableIndex = sheet.ListObjects.Add(
                firstRow,
                firstColumn,
                firstRow + totalRows - 1,
                firstColumn + totalColumns - 1,
                true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "SampleTable";
            table.ShowHeaderRow = true;
            // ShowTotalRow property is not available in this version; omitted.

            // Add a comment to cell A1
            int commentRow = 0;   // row 1
            int commentColumn = 0; // column A
            int commentIndex = sheet.Comments.Add(commentRow, commentColumn);
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "For more information, see documentation: https://example.com/docs";
            comment.IsVisible = true; // make the comment visible by default

            // Determine output file path
            string outputPath = "TableWithComment.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (string.IsNullOrEmpty(outputDir))
            {
                outputDir = Directory.GetCurrentDirectory();
            }
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
