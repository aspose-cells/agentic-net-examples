// Title: Add a visible comment to a header cell and freeze the header row using Aspose.Cells for .NET (C#)
// AI Prompts: Create a new workbook, add a comment with text and author to cell A1, set the comment to visible, then freeze row 1 by calling FreezePanes at cell A2 in C# with Aspose.Cells. | Generate an Excel file where the top‑row header contains a persistent note by inserting a comment into A1 and applying FreezePanes to keep the row fixed while scrolling. | Write C# code that adds a header comment, makes it always displayed, and freezes the first worksheet row using the Aspose.Cells FreezePanes method.
// Common Searches: how to add a visible comment to a header cell and freeze the row in Aspose.Cells C# | Aspose.Cells FreezePanes after inserting a comment in Excel workbook | C# code to keep Excel header comment on screen while scrolling | set comment author and make it visible then freeze top row using Aspose.Cells | freeze first row of worksheet after adding a comment with Aspose.Cells for .NET
// Tags: add comment to cell A1 Aspose.Cells | set comment visibility C# | freeze top worksheet row using FreezePanes | header note persistence Excel Aspose | comment author property Aspose.Cells | freeze panes after comment insertion

using System;
using Aspose.Cells;

namespace AsposeCellsCommentFreezeDemo
{
    // Demonstrates adding a visible comment with author to cell A1, making it always shown, freezing the first row via FreezePanes at A2, and saving the workbook as HeaderCommentFreeze.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a comment to the header cell (A1)
            CommentCollection comments = worksheet.Comments;
            int commentIndex = comments.Add("A1");               // Add comment to cell A1
            Comment headerComment = comments[commentIndex];      // Retrieve the comment object
            headerComment.Note = "This is a header note.";       // Set comment text
            headerComment.Author = "Demo Author";                // Optional: set author
            headerComment.IsVisible = true;                      // Make the comment visible

            // Freeze the row that contains the comment (row 1) so it stays visible while scrolling
            // FreezePanes(cellName, freezedRows, freezedColumns)
            // Freezing at "A2" with 1 frozen row keeps the first row (where the comment is) fixed.
            worksheet.FreezePanes("A2", 1, 0);

            // Save the workbook
            workbook.Save("HeaderCommentFreeze.xlsx");
        }
    }
}
