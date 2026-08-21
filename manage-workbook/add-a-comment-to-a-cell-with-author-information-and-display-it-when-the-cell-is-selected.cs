// Title: Add an Author Comment to a Cell and Show It on Selection with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, insert a comment with author and text into cell A1, set the comment to appear when the cell is selected, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells comment author | C# add cell comment | show comment on selection | comment visibility Aspose.Cells | Aspose.Cells workbook annotation | .NET spreadsheet comment
// Common Searches: Aspose.Cells add comment with author C# | display cell comment on selection .NET | make comment visible when cell is selected Aspose.Cells | set comment author in Excel file using C# | how to show Excel comment on click with Aspose
// Developer Intent: Insert a comment that includes author information into a specific cell and configure it to become visible only when the cell is selected.
// Use Cases: Document reviewer notes that appear on demand for audit trails. | Interactive spreadsheet guides where each cell reveals its author's comment on click. | Change‑tracking system that tags modified cells with author‑identified comments displayed upon selection.
// AI Prompts: Generate C# code using Aspose.Cells to add an author comment to cell B2 and make it visible only when the cell is selected. | Provide an example that loops through a range and adds author‑tagged comments to each cell, setting IsVisible to true for every comment. | Explain how to programmatically toggle comment visibility based on cell selection events in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, insert a comment with author and text into cell A1, set the comment to appear when the cell is selected, and save the file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a comment to cell A1
        int commentIndex = worksheet.Comments.Add("A1");
        Comment comment = worksheet.Comments[commentIndex];

        // Set author information and comment text
        comment.Author = "John Doe";
        comment.Note = "Reviewed by John Doe.";

        // Make the comment visible when the cell is selected
        comment.IsVisible = true;

        // Save the workbook
        workbook.Save("CommentWithAuthor.xlsx");
    }
}
