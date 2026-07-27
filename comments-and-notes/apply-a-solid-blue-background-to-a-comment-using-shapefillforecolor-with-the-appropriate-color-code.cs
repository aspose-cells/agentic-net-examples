// Title: Set a Solid Blue Background for a Worksheet Comment Using Shape.FillFormat.ForeColor in Aspose.Cells (C#)
// Description: Creates a new workbook, adds a comment to cell A1, accesses the comment's Shape, enables filling, sets FillFormat.ForeColor to a solid blue color, and saves the file, resulting in a comment with a blue background.
// Keywords: Aspose.Cells comment background | C# comment shape fill | Shape.FillFormat.ForeColor | solid blue comment color | Excel comment formatting Aspose | Aspose.Cells for .NET | set comment fill color | worksheet comment appearance
// Common Searches: Aspose.Cells set comment background color C# | How to fill comment shape with blue in Aspose.Cells | Change Excel comment fill using Shape.FillFormat | Apply solid color to comment in Aspose.Cells .NET | Set ForeColor for comment shape Aspose
// Developer Intent: Apply a solid blue fill to a worksheet comment via its Shape object.
// Use Cases: Visually highlight critical notes by giving comments a distinct blue background. | Enforce corporate branding on comment colors across all worksheets. | Create easily recognizable review comments for collaborative editing.
// AI Prompts: Write C# code with Aspose.Cells to set a custom RGB background for a comment shape. | Show how to apply a gradient fill to a comment instead of a solid color. | Explain how to retrieve an existing comment's Shape and modify its FillFormat properties.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsCommentBlueBackground
{
    // Creates a new workbook, adds a comment to cell A1, accesses the comment's Shape, enables filling, sets FillFormat.ForeColor to a solid blue color, and saves the file, resulting in a comment with a blue background.
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
            comment.Note = "This comment has a solid blue background.";

            // Access the shape attached to the comment
            Shape commentShape = comment.CommentShape;

            // Ensure the shape's fill is visible
            commentShape.IsFilled = true;

            // Apply a solid blue background using the FillFormat's ForeColor property
            commentShape.FillFormat.ForeColor = Color.Blue;

            // Save the workbook
            workbook.Save("CommentBlueBackground.xlsx");
        }
    }
}
