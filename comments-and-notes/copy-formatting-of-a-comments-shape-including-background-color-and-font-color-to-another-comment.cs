// Title: Copy Comment Shape Fill and Font Color Between Cells with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add formatted comments to cells, and transfer the comment shape's fill type, solid fill color, and font color from a source comment (A1) to a destination comment (B2) using Aspose.Cells for .NET, then save the file.
// Keywords: Aspose.Cells | C# | CommentShape | copy comment formatting | comment background color | comment font color | fill type | solid fill | Excel comment style | Aspose.Cells API
// Common Searches: Aspose.Cells copy comment shape formatting | C# copy Excel comment background and font color | transfer comment shape fill type with Aspose.Cells | duplicate comment visual style programmatically
// Developer Intent: Copy the fill type, solid fill color, and font color from one comment's shape to another comment's shape.
// Use Cases: Apply a consistent visual style to multiple comments in generated reports. | Synchronize comment appearance after bulk creation of comments across a worksheet. | Preserve comment formatting when copying data between workbooks.
// AI Prompts: Generate a reusable C# method that copies all shape formatting (fill, line, font) from a source CommentShape to any target CommentShape using Aspose.Cells. | Show how to copy gradient fills and font styles from one comment shape to multiple destinations in a workbook. | Explain step‑by‑step how to clone comment shape properties, including border and alignment, with Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, add formatted comments to cells, and transfer the comment shape's fill type, solid fill color, and font color from a source comment (A1) to a destination comment (B2) using Aspose.Cells for .NET, then save the file.
class CopyCommentShapeFormatting
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // ---------- Source comment (with formatting) ----------
        // Add a comment to cell A1
        int srcIdx = sheet.Comments.Add("A1");
        Comment srcComment = sheet.Comments[srcIdx];
        srcComment.Note = "Source comment";

        // Access the shape attached to the source comment
        CommentShape srcShape = srcComment.CommentShape;

        // Set background (fill) color
        srcShape.Fill.FillType = FillType.Solid;
        srcShape.Fill.SolidFill.Color = Color.LightYellow;

        // Set font color of the shape's text
        srcShape.Font.Color = Color.Blue;

        // ---------- Destination comment ----------
        // Add a comment to cell B2 (the one that will receive the formatting)
        int destIdx = sheet.Comments.Add("B2");
        Comment destComment = sheet.Comments[destIdx];
        destComment.Note = "Destination comment";

        // Access the shape attached to the destination comment
        CommentShape destShape = destComment.CommentShape;

        // ----- Copy formatting from source shape to destination shape -----
        // Copy fill type
        destShape.Fill.FillType = srcShape.Fill.FillType;

        // If the fill is solid, copy the solid fill color
        if (srcShape.Fill.FillType == FillType.Solid)
        {
            destShape.Fill.SolidFill.Color = srcShape.Fill.SolidFill.Color;
        }

        // Copy font color
        destShape.Font.Color = srcShape.Font.Color;

        // Save the workbook
        workbook.Save("CommentShapeFormattingCopy.xlsx");
    }
}
