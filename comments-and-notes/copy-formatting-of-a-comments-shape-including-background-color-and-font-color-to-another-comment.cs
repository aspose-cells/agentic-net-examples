// Title: Copy comment shape background and font colors from one cell to another with Aspose.Cells for .NET
// AI Prompts: Write C# code that copies the fill color and font color of a source comment's CommentShape to a target comment's CommentShape in an Aspose.Cells workbook. | Generate a method that transfers comment shape formatting (background and text color) between two cells using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells copy comment shape fill color from cell A1 to B2 C# | How to duplicate comment background and font color in Excel using Aspose.Cells .NET | C# Aspose.Cells transfer comment shape formatting between cells | Copy comment style from one worksheet cell to another Aspose.Cells | Set comment shape fill and font colors programmatically with Aspose.Cells
// Tags: copy comment shape formatting Aspose.Cells | comment shape fill color Aspose.Cells | comment shape font color Aspose.Cells | Aspose.Cells comment style replication | C# transfer comment formatting Excel

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // This example creates a workbook, adds a source comment with a light‑yellow fill and blue font, adds a destination comment, copies the fill and font colors from the source comment's CommentShape to the destination comment's CommentShape, and saves the file.
class CopyCommentShapeFormatting
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // ---------- Source comment ----------
        // Add a comment to cell A1
        int srcCommentIndex = worksheet.Comments.Add("A1");
        Comment srcComment = worksheet.Comments[srcCommentIndex];
        srcComment.Note = "Source comment";

        // Get the shape attached to the source comment
        CommentShape srcShape = srcComment.CommentShape;

        // Set background (fill) color and font color for the source shape
        srcShape.FillFormat.ForeColor = Color.LightYellow;   // background color
        srcShape.Font.Color = Color.Blue;                    // font color

        // ---------- Destination comment ----------
        // Add a comment to cell B2
        int destCommentIndex = worksheet.Comments.Add("B2");
        Comment destComment = worksheet.Comments[destCommentIndex];
        destComment.Note = "Destination comment";

        // Get the shape attached to the destination comment
        CommentShape destShape = destComment.CommentShape;

        // Copy formatting from source shape to destination shape
        destShape.FillFormat.ForeColor = srcShape.FillFormat.ForeColor; // background color
        destShape.Font.Color = srcShape.Font.Color;                    // font color

        // Save the workbook
        workbook.Save("CommentShapeFormattingCopy.xlsx");
    }
}
