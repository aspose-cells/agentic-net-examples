// Title: Aspose.Cells for .NET – Set Comment Text Orientation to TopToBottom (Vertical)
// Description: Creates a workbook, adds a comment to cell B2, changes the comment's TextOrientationType to TopToBottom for stacked vertical text, resizes the comment shape, and saves the file as CommentTopToBottom.xlsx.
// Keywords: Aspose.Cells comment orientation | TopToBottom text direction | vertical comment Aspose.Cells | C# set comment text orientation | Excel comment shape size
// Common Searches: Aspose.Cells set comment vertical text | C# comment TextOrientationType TopToBottom example | how to make Excel comment display vertically | adjust comment width height for vertical text Aspose
// Developer Intent: Apply a vertical (TopToBottom) text layout to a comment shape in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Label column headers with vertical comments in automated reports. | Improve readability of notes in narrow columns by stacking text vertically. | Support top‑to‑bottom scripts or design requirements in generated spreadsheets.
// AI Prompts: Generate C# code with Aspose.Cells that adds a comment to a cell and sets its TextOrientationType to TopToBottom. | Explain how to resize a comment shape after changing its orientation to ensure the vertical text is fully visible. | Provide a loop that iterates through all comments in a worksheet and changes each one to a TopToBottom orientation.

using System;
using Aspose.Cells;

// Creates a workbook, adds a comment to cell B2, changes the comment's TextOrientationType to TopToBottom for stacked vertical text, resizes the comment shape, and saves the file as CommentTopToBottom.xlsx.
class SetCommentTextDirection
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a comment to cell B2
        int commentIndex = worksheet.Comments.Add("B2");
        Comment comment = worksheet.Comments[commentIndex];
        comment.Note = "Vertical comment text";

        // Set the comment's text orientation to TopToBottom (stacked vertical layout)
        comment.TextOrientationType = TextOrientationType.TopToBottom;

        // Adjust comment size so the vertical text is visible
        comment.Width = 150;
        comment.Height = 100;

        // Save the workbook
        workbook.Save("CommentTopToBottom.xlsx");
    }
}
