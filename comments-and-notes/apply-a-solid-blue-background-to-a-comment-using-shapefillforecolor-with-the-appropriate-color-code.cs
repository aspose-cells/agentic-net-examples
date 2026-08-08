// Title: Aspose.Cells for .NET – Apply a Solid Blue Background to a Worksheet Comment via Shape.Fill.ForeColor
// Description: Creates a workbook, adds a comment to cell A1, retrieves the comment's Shape, enables filling, and sets FillFormat.ForeColor to a blue color, producing a solid blue background before saving the file as an .xlsx document.
// Keywords: Aspose.Cells comment background color | Shape.Fill.ForeColor .NET | solid blue comment fill | worksheet comment styling | Aspose.Cells FillFormat example
// Common Searches: how to change comment background color in Aspose.Cells | set blue fill for comment shape Aspose.Cells .NET | apply solid fill to worksheet comment using Shape.Fill.ForeColor | Aspose.Cells comment shape color customization
// Developer Intent: Add a solid blue fill to a worksheet comment using the Shape.Fill.ForeColor property.
// Use Cases: Brand report cells with a uniform blue comment background for visual consistency. | Highlight critical notes in generated spreadsheets by applying a distinct color to comments. | Create templates where comments automatically match corporate color guidelines.
// AI Prompts: Generate code to set a custom RGB value for a comment's background using Shape.Fill.ForeColor in Aspose.Cells for .NET. | Show how to apply a gradient fill to a comment shape instead of a solid color. | Explain how to hide or show a comment's fill without altering its text content.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, adds a comment to cell A1, retrieves the comment's Shape, enables filling, and sets FillFormat.ForeColor to a blue color, producing a solid blue background before saving the file as an .xlsx document.
class ApplyBlueBackgroundToComment
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a comment to cell A1
        int commentIndex = worksheet.Comments.Add("A1");
        Comment comment = worksheet.Comments[commentIndex];
        comment.Note = "This comment has a solid blue background.";

        // Access the shape attached to the comment
        Shape commentShape = comment.CommentShape;

        // Ensure the shape's fill is visible
        commentShape.IsFilled = true;

        // Set the fill fore color to solid blue using the FillFormat (obsolete) property
        // This applies a solid background color to the comment shape
        commentShape.FillFormat.ForeColor = Color.Blue;

        // Save the workbook
        workbook.Save("CommentWithBlueBackground.xlsx");
    }
}
