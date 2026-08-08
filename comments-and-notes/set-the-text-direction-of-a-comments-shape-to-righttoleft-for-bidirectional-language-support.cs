// Title: C# – Set Comment Shape Text Direction to RightToLeft with Aspose.Cells
// Description: Shows how to create a workbook, add a comment, access its CommentShape, set the TextDirection property to RightToLeft for bidirectional language rendering, and save the Excel file.
// Keywords: Aspose.Cells | C# | CommentShape | TextDirection | RightToLeft | bidirectional language | Excel comment | Arabic | Hebrew
// Common Searches: Aspose.Cells set comment shape RightToLeft | C# change comment text direction in Excel | RightToLeft comment rendering Aspose.Cells | How to enable Arabic comments in Excel with Aspose.Cells | Set TextDirection of CommentShape using Aspose.Cells .NET
// Developer Intent: Apply RightToLeft text direction to a comment's shape to support Arabic, Hebrew, or other bidirectional scripts.
// Use Cases: Generate reports where comment boxes must display Arabic or Hebrew text correctly. | Update existing workbooks to convert comment orientation to RightToLeft before distribution. | Create multilingual Excel templates with comments that automatically adapt to right‑to‑left languages.
// AI Prompts: Provide C# code that sets a comment's shape TextDirection to RightToLeft using Aspose.Cells. | How can I enable right‑to‑left text in Excel comment shapes for Arabic content with Aspose.Cells? | Explain the effect of TextDirectionType.RightToLeft on comment rendering in an Aspose.Cells workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsCommentShapeDirection
{
    // Shows how to create a workbook, add a comment, access its CommentShape, set the TextDirection property to RightToLeft for bidirectional language rendering, and save the Excel file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a comment to cell A1
            int commentIndex = worksheet.Comments.Add("A1");
            Comment comment = worksheet.Comments[commentIndex];
            comment.Note = "This comment's shape will have RightToLeft text direction.";

            // Obtain the shape associated with the comment
            CommentShape commentShape = comment.CommentShape;

            // Set the text direction of the comment's shape to RightToLeft
            commentShape.TextDirection = TextDirectionType.RightToLeft;

            // Save the workbook (lifecycle save rule)
            workbook.Save("CommentShapeTextDirection.xlsx");
        }
    }
}
