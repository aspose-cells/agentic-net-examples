// Title: Aspose.Cells C# – Insert Rectangle Shape with Hover Tooltip Comment
// Description: Shows how to create a workbook, add a rectangle shape at cell B2, set its AlternativeText to display a tooltip on hover, attach a comment to the same cell, customize the CommentShape size and colors, and save the result as an .xlsx file.
// Keywords: Aspose.Cells C# shape tooltip | add rectangle shape Aspose.Cells | Excel hover comment Aspose | AlternativeText shape tooltip | CommentShape customization | C# generate Excel tooltip | Aspose.Cells shape annotation
// Common Searches: Aspose.Cells add shape with tooltip C# | Set AlternativeText for rectangle shape Aspose.Cells | Customize comment box appearance Aspose.Cells | Display hover tooltip for Excel shape using Aspose | How to attach comment to shape in Aspose.Cells
// Developer Intent: Create a rectangle shape and bind a comment that appears as a tooltip when the user hovers over the shape in an Excel workbook.
// Use Cases: Add interactive notes to financial dashboards by linking shapes with hover tooltips. | Provide explanatory tooltips for diagram elements in automatically generated reports. | Style comment boxes attached to shapes for a polished UI in exported Excel files.
// AI Prompts: Generate C# code with Aspose.Cells that adds a circle shape and a hover tooltip comment, including custom fill and line colors. | Write a method that inserts multiple shapes, each with its own tooltip comment, and saves the workbook. | Explain how the AlternativeText property of a Shape and the CommentShape of a Comment work together to display hover tooltips in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to create a workbook, add a rectangle shape at cell B2, set its AlternativeText to display a tooltip on hover, attach a comment to the same cell, customize the CommentShape size and colors, and save the result as an .xlsx file.
class ShapeWithCommentTooltip
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet (positioned at row 2, column 2)
        // Parameters: upper left row, upper left column, upper left row offset, upper left column offset, width, height
        Shape rectShape = sheet.Shapes.AddRectangle(1, 1, 0, 0, 120, 60);
        rectShape.Name = "MyRectangle";
        rectShape.AlternativeText = "Rectangle shape tooltip";

        // Add a comment to cell B2 (row index 1, column index 1)
        int commentIndex = sheet.Comments.Add(1, 1);
        Comment comment = sheet.Comments[commentIndex];
        comment.Note = "This is a tooltip comment displayed on hover.";

        // Access the shape that represents the comment box
        CommentShape commentShape = comment.CommentShape;
        // Optionally adjust the comment shape size and appearance
        commentShape.Width = 200;
        commentShape.Height = 80;
        commentShape.FillFormat.ForeColor = System.Drawing.Color.LightYellow;
        commentShape.LineFormat.ForeColor = System.Drawing.Color.DarkGray;

        // Save the workbook to a file
        workbook.Save("ShapeWithCommentTooltip.xlsx");
    }
}
