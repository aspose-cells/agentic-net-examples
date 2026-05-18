using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsCommentFormattingCopy
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the comments collection of the worksheet
            CommentCollection comments = worksheet.Comments;

            // -------------------------------------------------
            // Create source comment (A1) and set its shape format
            // -------------------------------------------------
            int srcIndex = comments.Add("A1");
            Comment srcComment = comments[srcIndex];
            srcComment.Note = "Source comment";

            // Set background color (fill) of the source comment shape
            CommentShape srcShape = srcComment.CommentShape;
            srcShape.Fill.FillType = FillType.Solid;
            srcShape.Fill.SolidFill.Color = Color.LightYellow;   // background color

            // Set font color of the source comment shape
            srcShape.Font.Color = Color.Blue;                    // font color

            // -------------------------------------------------
            // Create destination comment (B2) – formatting will be copied
            // -------------------------------------------------
            int destIndex = comments.Add("B2");
            Comment destComment = comments[destIndex];
            destComment.Note = "Destination comment";

            // Get the shape of the destination comment
            CommentShape destShape = destComment.CommentShape;

            // -------------------------------
            // Copy formatting from source to destination
            // -------------------------------
            // Copy background (fill) settings
            destShape.Fill.FillType = srcShape.Fill.FillType;
            destShape.Fill.SolidFill.Color = srcShape.Fill.SolidFill.Color;

            // Copy font color
            destShape.Font.Color = srcShape.Font.Color;

            // Save the workbook to verify the result
            workbook.Save("CommentFormattingCopy.xlsx");
        }
    }
}