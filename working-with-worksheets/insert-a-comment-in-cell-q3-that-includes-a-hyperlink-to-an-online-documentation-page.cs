// Title: Insert a visible comment containing a documentation hyperlink into cell Q3 with Aspose.Cells for .NET
// AI Prompts: Create a comment in cell Q3 that shows a clickable link to the Aspose.Cells .NET documentation and set the comment to be visible. | Update the workbook code to add a visible comment with the documentation URL in column Q row 3 and then save the file.
// Common Searches: Aspose.Cells .NET how to add a comment with a URL to cell Q3 | make Excel comment visible and include hyperlink using Aspose.Cells | insert documentation link inside a comment in an Aspose.Cells workbook | add visible comment to specific cell in Excel with Aspose.Cells API
// Tags: add comment with hyperlink Aspose.Cells .NET | visible comment in Excel workbook | insert comment into cell Q3 | hyperlink inside comment Aspose.Cells | set comment visibility Aspose.Cells

using System;
using Aspose.Cells;

// The program creates a new workbook, adds a visible comment that contains the Aspose.Cells .NET documentation URL to cell Q3 (row 3, column Q), and saves the file as CommentWithHyperlink.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Target cell Q3 (zero‑based indices)
            int row = 2;      // Row 3
            int column = 16;  // Column Q

            // Add a comment to the cell
            int commentIndex = sheet.Comments.Add(row, column);
            Comment comment = sheet.Comments[commentIndex];

            // Set comment text (hyperlink as plain text for compatibility)
            comment.Note = "Aspose.Cells Documentation\nhttps://docs.aspose.com/cells/net/";

            // Make the comment visible
            comment.IsVisible = true;

            // Save the workbook
            workbook.Save("CommentWithHyperlink.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
