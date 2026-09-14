// Title: Create a styled ListObject table and insert a SharePoint hyperlink comment in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that creates a worksheet, defines a range, converts it to a ListObject table with headers, applies a medium table style, and adds a comment containing a SharePoint URL to cell A1. | Write a C# example that saves the workbook to a specified file path, automatically creating the output directory if it does not exist, using Aspose.Cells. | Show how to add a plain‑text comment with a clickable SharePoint link to an Excel cell via Aspose.Cells.
// Common Searches: asp.net aspose.cells create excel ListObject table with header and style | c# add comment with hyperlink to excel cell using Aspose.Cells | asp.net save workbook to folder creating directory if missing Aspose.Cells | how to apply TableStyleMedium9 to a table in Aspose.Cells C# | insert SharePoint document link in Excel cell comment via Aspose.Cells
// Tags: Aspose.Cells ListObject table creation C# | Aspose.Cells apply table style .xlsx | Aspose.Cells add cell comment hyperlink | Aspose.Cells ensure output directory exists | Aspose.Cells save workbook with comment

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The example creates a new workbook, defines a two‑column range, converts it into a ListObject table with headers, applies the TableStyleMedium9 style, adds a comment to cell A1 that contains a SharePoint document URL, ensures the target directory exists, and saves the file as TableWithComment.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data that will become the table
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");

            // Define the range that will be converted into a table (ListObject)
            CellArea tableArea = new CellArea
            {
                StartRow = 0,   // A1 row index
                StartColumn = 0,
                EndRow = 2,     // A3 row index
                EndColumn = 1
            };

            // Add the table to the worksheet (hasHeaders = true because first row contains column names)
            int tableIndex = sheet.ListObjects.Add(
                tableArea.StartRow,
                tableArea.StartColumn,
                tableArea.EndRow,
                tableArea.EndColumn,
                true);

            ListObject table = sheet.ListObjects[tableIndex];
            table.ShowHeaderRow = true;
            table.TableStyleType = TableStyleType.TableStyleMedium9;

            // Add a comment to cell A1 that contains a hyperlink to a SharePoint document
            // The hyperlink appears as plain text; Excel will treat it as a clickable link.
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "Reference document: https://sharepoint.company.com/sites/docs/Reference.docx";

            // Save the workbook to a file
            string outputPath = "TableWithComment.xlsx";
            // Ensure the directory exists (in case a relative path is used)
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
