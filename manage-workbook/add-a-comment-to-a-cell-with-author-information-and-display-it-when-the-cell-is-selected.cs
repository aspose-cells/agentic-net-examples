// Title: Add a comment with author and make it visible on cell selection using Aspose.Cells for .NET (C#)
// Common Searches: how to add a comment with author to a specific cell in Aspose.Cells C# | display comment automatically when a cell is selected Aspose.Cells .NET | set comment author and visibility flag for Excel cell using Aspose.Cells | Aspose.Cells add comment to A1 and show on selection C# example | make Excel cell comment visible on click with Aspose.Cells API
// Tags: cell comment author Aspose.Cells C# | comment visibility on selection Aspose.Cells | set comment IsVisible property .NET | programmatic Excel comment creation Aspose.Cells | add comment to A1 worksheet Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new workbook, writes a value to cell A1, adds a comment to that cell, sets the comment text and author, configures the comment to be visible when the cell is selected, and saves the file as CommentWithAuthor.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle create rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Put a sample value into cell A1
        Cell cell = sheet.Cells["A1"];
        cell.PutValue("Sample Data");

        // Add a comment to cell A1
        // The Add method returns the index of the new comment
        int commentIndex = sheet.Comments.Add("A1");
        Comment comment = sheet.Comments[commentIndex];

        // Set the comment text
        comment.Note = "Reviewed by John Doe on 2026-11-23.";

        // Set the author of the comment
        comment.Author = "John Doe";

        // Make the comment visible when the cell is selected
        comment.IsVisible = true;

        // Save the workbook (lifecycle save rule)
        workbook.Save("CommentWithAuthor.xlsx");
    }
}
