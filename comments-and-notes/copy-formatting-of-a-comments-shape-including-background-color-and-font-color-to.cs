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

            // -------------------------------------------------
            // Add source comment to cell A1 and set its formatting
            // -------------------------------------------------
            int srcIndex = worksheet.Comments.Add("A1");
            Comment srcComment = worksheet.Comments[srcIndex];
            srcComment.Note = "Source comment";

            // Set font color of the source comment
            srcComment.Font.Color = Color.Blue;

            // Access the shape attached to the source comment
            Shape srcShape = srcComment.CommentShape;

            // Set background (fill) color of the source comment shape
            srcShape.Fill.FillType = FillType.Solid;
            srcShape.Fill.SolidFill.Color = Color.LightYellow;

            // -------------------------------------------------
            // Add destination comment to cell B2 (initially with default formatting)
            // -------------------------------------------------
            int destIndex = worksheet.Comments.Add("B2");
            Comment destComment = worksheet.Comments[destIndex];
            destComment.Note = "Destination comment";

            // Access the shape attached to the destination comment
            Shape destShape = destComment.CommentShape;

            // -------------------------------------------------
            // Copy formatting from source comment to destination comment
            // -------------------------------------------------
            // Copy font color
            destComment.Font.Color = srcComment.Font.Color;

            // Copy shape background color
            destShape.Fill.FillType = srcShape.Fill.FillType;
            if (srcShape.Fill.FillType == FillType.Solid)
            {
                destShape.Fill.SolidFill.Color = srcShape.Fill.SolidFill.Color;
            }

            // Copy line color (use ForeColor property)
            destShape.Line.ForeColor = srcShape.Line.ForeColor;

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            workbook.Save("CommentFormattingCopy.xlsx");
        }
    }
}