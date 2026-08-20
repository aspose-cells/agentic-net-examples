// Title: Add a visible comment to a header cell and freeze the top row with Aspose.Cells for .NET
// Description: Shows how to create a workbook in C#, write a header in A1, attach a visible comment (author, autosize), freeze the first row via FreezePanes, and save the result as HeaderCommentWithFreeze.xlsx.
// Keywords: Aspose.Cells comment visible | C# freeze panes | add comment to cell A1 | freeze top row programmatically | Aspose.Cells header note | .NET workbook comment author | auto‑size comment Aspose.Cells
// Common Searches: Aspose.Cells add visible comment to header | freeze first row after adding comment C# | keep cell comment visible while scrolling Aspose.Cells | set comment author and autosize in .NET | FreezePanes example with comments
// Developer Intent: Insert a comment on a header cell and ensure it stays visible by freezing the worksheet’s top row.
// Use Cases: Provide explanatory notes on column headers that remain on‑screen during scrolling. | Create data‑entry templates where header comments describe required formats. | Generate reports with frozen header rows that include persistent reviewer comments.
// AI Prompts: Generate C# code to add a visible comment to cell A1, set the author, enable autosize, and freeze the first row using Aspose.Cells. | Explain how to customize comment font, color, and background while keeping it visible after applying FreezePanes. | Show an example of adding comments to multiple header cells and freezing the top two rows in a .NET workbook.

using System;
using Aspose.Cells;

// Shows how to create a workbook in C#, write a header in A1, attach a visible comment (author, autosize), freeze the first row via FreezePanes, and save the result as HeaderCommentWithFreeze.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Put a header value in cell A1
        worksheet.Cells["A1"].PutValue("Header");

        // Add a comment to the header cell (A1)
        int commentIndex = worksheet.Comments.Add("A1");
        Comment comment = worksheet.Comments[commentIndex];
        comment.Note = "This is a header comment.";
        comment.Author = "Admin";
        comment.IsVisible = true;   // Make the comment visible by default
        comment.AutoSize = true;    // Adjust size automatically

        // Freeze the first row so the comment stays visible while scrolling
        // Freeze at cell A2 (row index 1) with 1 frozen row and 0 frozen columns
        worksheet.FreezePanes("A2", 1, 0);

        // Save the workbook
        workbook.Save("HeaderCommentWithFreeze.xlsx");
    }
}
