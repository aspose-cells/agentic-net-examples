// Title: Aspose.Cells for .NET – Add a Comment to Cell D4 in C#
// Description: Shows how to create a new workbook with Aspose.Cells, insert a comment containing "Review required before final submission." into cell D4, and save the file as CommentInD4.xlsx using C#.
// Keywords: Aspose.Cells | C# Excel comment | add comment D4 | Aspose.Cells API | Excel note .NET | Workbook.Save | Aspose.Cells tutorial
// Common Searches: Aspose.Cells add comment to specific cell | C# insert comment D4 Excel | How to set comment text with Aspose.Cells | Save workbook after adding comments Aspose.Cells | Aspose.Cells comment visibility settings
// Developer Intent: Insert a predefined comment into cell D4 of a newly created Excel workbook using Aspose.Cells for .NET.
// Use Cases: Mark cells that need review before publishing the workbook. | Provide reviewers with inline feedback directly in the spreadsheet. | Create an audit trail by programmatically attaching notes to key cells. | Automate quality‑control annotations during report generation.
// AI Prompts: Write C# code with Aspose.Cells to add a comment to cell D4 and set the author to "QA Team". | Show how to update an existing comment in cell D4, change its text, and make it visible only when the cell is selected. | Provide a loop example that adds comments to a range of cells (e.g., D4:F6) with custom messages using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Shows how to create a new workbook with Aspose.Cells, insert a comment containing "Review required before final submission." into cell D4, and save the file as CommentInD4.xlsx using C#.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a comment to cell D4
        int commentIndex = worksheet.Comments.Add("D4");
        Comment comment = worksheet.Comments[commentIndex];
        comment.Note = "Review required before final submission.";

        // Save the workbook
        workbook.Save("CommentInD4.xlsx");
    }
}
