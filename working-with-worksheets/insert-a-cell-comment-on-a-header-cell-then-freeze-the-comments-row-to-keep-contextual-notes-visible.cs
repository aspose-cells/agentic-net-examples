// Title: Insert a visible comment on a header cell and freeze its row using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add a visible comment with custom author and font to cell A1, freeze the first row so the comment stays in view while scrolling, and save the file as CommentWithFrozenHeader.xlsx.
// Keywords: Aspose.Cells add comment C# | freeze row Aspose.Cells | visible comment Aspose.Cells | freeze panes header Aspose.Cells | set comment author Aspose.Cells | C# workbook comment freeze | Aspose.Cells comment formatting
// Common Searches: Aspose.Cells add comment to cell A1 C# | How to keep a comment visible while scrolling in Aspose.Cells | Freeze header row with comment using Aspose.Cells .NET | Make comment always visible in Excel with Aspose.Cells | C# Aspose.Cells freeze panes based on comment row
// Developer Intent: Add a persistent comment to a header cell and lock its row in the worksheet.
// Use Cases: Provide on‑screen guidance for column headers in financial reports | Create a template where explanatory notes stay fixed for end‑users | Generate audit‑ready spreadsheets with author‑attributed comments that remain visible | Design dashboards where header comments are frozen for quick reference
// AI Prompts: Generate C# Aspose.Cells code that adds a comment to cell B2, sets the author to 'Reviewer', makes it visible, and freezes rows up to that comment. | Show how to add multiple header comments and freeze the top two rows in an Aspose.Cells workbook using C#. | Explain the FreezePanes parameters and how they affect rows containing comments in Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, add a visible comment with custom author and font to cell A1, freeze the first row so the comment stays in view while scrolling, and save the file as CommentWithFrozenHeader.xlsx.
class InsertCommentAndFreeze
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a comment to the header cell A1
        int commentIndex = worksheet.Comments.Add("A1");
        Comment comment = worksheet.Comments[commentIndex];
        comment.Note = "Header note: This column contains sales data.";
        comment.Author = "Analyst";
        comment.Font.Name = "Calibri";
        comment.Font.Size = 11;
        comment.IsVisible = true;

        // Freeze the row that contains the comment (row 0, first row)
        int commentRow = comment.Row; // zero‑based index
        // Freeze panes so that rows up to and including the comment row are locked
        worksheet.FreezePanes(commentRow + 1, 0, commentRow + 1, 0);

        // Save the workbook
        workbook.Save("CommentWithFrozenHeader.xlsx");
    }
}
