// Title: Apply a solid blue background to an Excel comment using Aspose.Cells Shape.FillFormat.ForeColor in C#
// AI Prompts: Use Shape.FillFormat.ForeColor to set a comment's fill to solid blue in an Aspose.Cells workbook (C#). | Create a comment on a cell and change its background color to blue via the comment's Shape properties with Aspose.Cells. | Save the workbook after applying a blue fill to the comment shape using Aspose.Cells C# API.
// Common Searches: Aspose.Cells C# set comment background color to blue | How to change Excel comment fill color with Shape.FillFormat.ForeColor | Apply solid fill to comment shape in Aspose.Cells workbook | C# code to make Excel comment have blue background using Aspose.Cells | Enable comment shape fill and set color in Aspose.Cells
// Tags: comment shape fill color Aspose.Cells | Shape.FillFormat.ForeColor C# | Excel comment solid background Aspose.Cells | set comment shape fill Aspose.Cells | blue fill comment Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, adds a comment to cell A1, enables the comment shape's fill, sets a solid blue background using Shape.FillFormat.ForeColor, and saves the file as CommentBlueBackground.xlsx.
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

        // Access the shape associated with the comment
        Shape commentShape = comment.CommentShape;

        // Ensure the shape's fill is visible
        commentShape.IsFilled = true;

        // Set a solid blue background using the (obsolete) FillFormat.ForeColor property
        commentShape.FillFormat.ForeColor = Color.Blue;

        // Save the workbook
        workbook.Save("CommentBlueBackground.xlsx");
    }
}
