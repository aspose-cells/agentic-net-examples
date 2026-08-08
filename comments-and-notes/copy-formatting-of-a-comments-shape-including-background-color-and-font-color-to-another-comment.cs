// Title: Copy comment shape fill and font formatting between cell comments with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a styled source comment in A1 (custom font and solid light‑yellow background), add a destination comment in B2, and transfer the source comment's shape fill type, solid fill color, and font attributes (color, name, size) to the destination comment before saving the file.
// Keywords: Aspose.Cells | .NET | C# | Excel comment shape | copy comment formatting | background fill color | font color | fill type | solid fill | programmatic comment style | cell comment styling
// Common Searches: Aspose.Cells copy comment background color | C# copy font color from one comment to another | transfer comment shape fill type Aspose.Cells | duplicate Excel comment formatting programmatically | how to copy comment style between cells using Aspose
// Developer Intent: Programmatically duplicate the visual style of a source comment's shape onto another comment in the same worksheet.
// Use Cases: Apply a predefined comment appearance across many cells in a template workbook. | Generate reports that automatically use a consistent comment look for newly added notes. | Enforce corporate comment styling by copying font and background settings from a master comment to all other comments.
// AI Prompts: Generate C# code with Aspose.Cells that copies the fill type and solid fill color from a source comment's CommentShape to a target comment's CommentShape. | Create a reusable method that accepts two cell addresses and copies all font and shape formatting from the source comment to the destination comment using Aspose.Cells for .NET. | Explain how to safely copy comment shape formatting when the source fill type may be non‑solid, handling both solid and pattern fills.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsCommentShapeCopyFormatting
{
    // Demonstrates how to create a workbook, add a styled source comment in A1 (custom font and solid light‑yellow background), add a destination comment in B2, and transfer the source comment's shape fill type, solid fill color, and font attributes (color, name, size) to the destination comment before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // Add two comments: source (A1) and destination (B2)
            // ------------------------------------------------------------
            // Source comment
            int srcIndex = worksheet.Comments.Add("A1");
            Comment srcComment = worksheet.Comments[srcIndex];
            srcComment.Note = "Source comment";
            srcComment.Font.Color = Color.Blue;               // Font color to copy
            srcComment.Font.Name = "Calibri";
            srcComment.Font.Size = 12;

            // Modify the shape of the source comment (background color)
            CommentShape srcShape = srcComment.CommentShape;
            srcShape.Fill.FillType = FillType.Solid;
            srcShape.Fill.SolidFill.Color = Color.LightYellow; // Background color to copy

            // Destination comment
            int destIndex = worksheet.Comments.Add("B2");
            Comment destComment = worksheet.Comments[destIndex];
            destComment.Note = "Destination comment";

            // ------------------------------------------------------------
            // Copy formatting from source comment's shape to destination comment's shape
            // ------------------------------------------------------------
            // Get the shape objects
            CommentShape destShape = destComment.CommentShape;

            // Copy background (fill) color
            destShape.Fill.FillType = srcShape.Fill.FillType;
            // Ensure the source shape uses Solid fill before accessing SolidFill.Color
            if (srcShape.Fill.FillType == FillType.Solid)
            {
                destShape.Fill.SolidFill.Color = srcShape.Fill.SolidFill.Color;
            }

            // Copy font color (and optionally other font properties)
            destShape.Font.Color = srcComment.Font.Color;
            destShape.Font.Name = srcComment.Font.Name;
            destShape.Font.Size = srcComment.Font.Size;

            // ------------------------------------------------------------
            // Save the workbook
            // ------------------------------------------------------------
            workbook.Save("CommentShapeFormattingCopied.xlsx");
        }
    }
}
