// Title: How to set a comment shape’s text orientation to TopToBottom for vertical comments in Excel using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a workbook, adds a comment to a cell, and sets the comment’s CommentShape.TextOrientationType to TopToBottom with Aspose.Cells. | Show the steps to configure a comment shape to display its text vertically (TopToBottom) before saving the file in Aspose.Cells for .NET. | Provide a minimal example that demonstrates changing the text direction of an Excel comment shape to vertical using the Aspose.Cells C# API.
// Common Searches: Aspose.Cells C# set comment shape text orientation to TopToBottom | vertical comment text in Excel using Aspose.Cells .NET | how to change comment text direction in Aspose.Cells workbook | C# Aspose.Cells comment shape TextOrientationType example | make Excel comment display vertically with Aspose.Cells API
// Tags: Aspose.Cells comment shape text orientation | C# set comment text direction TopToBottom | vertical comment annotation Aspose.Cells | Excel comment shape TextOrientationType .NET | Aspose.Cells comment shape properties

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a comment to cell A1, accesses its CommentShape, sets TextOrientationType to TopToBottom for vertical text layout, and saves the file as CommentShapeTopToBottom.xlsx.
class SetCommentShapeTextDirection
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a comment to cell A1
        int commentIdx = worksheet.Comments.Add("A1");
        Comment comment = worksheet.Comments[commentIdx];
        comment.Note = "This comment will be displayed vertically.";

        // Access the shape associated with the comment
        CommentShape commentShape = comment.CommentShape;

        // Set the text orientation of the comment shape to TopToBottom
        commentShape.TextOrientationType = TextOrientationType.TopToBottom;

        // Save the workbook
        workbook.Save("CommentShapeTopToBottom.xlsx");
    }
}
