// Title: Set comment shape text direction to TopToBottom (vertical) with Aspose.Cells for .NET
// Description: Shows how to add a comment to cell A1, retrieve its CommentShape, set TextOrientationType to TopToBottom for vertical text, and save the workbook as CommentTopToBottom.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | CommentShape | TextOrientationType | TopToBottom | vertical comment | C# Excel API | set comment orientation | Excel comment vertical text
// Common Searches: Aspose.Cells set comment vertical text | C# change comment orientation to TopToBottom | how to make Excel comment display top to bottom using Aspose | CommentShape TextOrientationType example | vertical comment layout Aspose.Cells .NET
// Developer Intent: Change a comment's text orientation to vertical (TopToBottom) in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Design worksheets where narrow columns require comments to be read top‑to‑bottom. | Create reports that need vertical annotations for better visual hierarchy. | Adjust dashboard comment layout to avoid overlapping adjacent cells.
// AI Prompts: Provide C# code with Aspose.Cells that sets a comment's TextOrientationType to TopToBottom. | Explain how to access a comment's CommentShape and modify its text direction in Aspose.Cells for .NET. | Show an example of saving a workbook after changing a comment's orientation to vertical.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsCommentTextDirection
{
    // Shows how to add a comment to cell A1, retrieve its CommentShape, set TextOrientationType to TopToBottom for vertical text, and save the workbook as CommentTopToBottom.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a comment to cell A1
            int commentIndex = worksheet.Comments.Add("A1");
            Comment comment = worksheet.Comments[commentIndex];
            comment.Note = "This comment will be displayed vertically (TopToBottom).";

            // Access the shape that represents the comment
            // CommentShape inherits from Shape and provides TextOrientationType property
            CommentShape commentShape = comment.CommentShape;

            // Set the text orientation of the comment shape to TopToBottom
            commentShape.TextOrientationType = TextOrientationType.TopToBottom;

            // Save the workbook
            workbook.Save("CommentTopToBottom.xlsx");
        }
    }
}
