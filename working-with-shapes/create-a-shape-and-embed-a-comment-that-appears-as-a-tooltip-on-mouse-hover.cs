// Title: Add a rectangle shape with a hover tooltip comment to an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add rectangle shape with tooltip on mouse hover | set alternative text for Excel shape using Aspose.Cells .NET | how to display a hover comment for a shape in an Excel file with Aspose.Cells | create shape tooltip in Excel workbook programmatically with Aspose.Cells | Aspose.Cells example adding shape and tooltip comment in C#
// Tags: add rectangle shape Aspose.Cells | shape alternative text tooltip Aspose.Cells | shape fill format Aspose.Cells | add cell comment Aspose.Cells | save workbook xlsx Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a new workbook, insert a rectangle shape, set its fill color, assign AlternativeText to show a tooltip on mouse hover, add a regular cell comment, and save the file as ShapeWithTooltip.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook.
        Workbook workbook = new Workbook();

        // Access the first worksheet.
        Worksheet sheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet.
        // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width.
        Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 80, 120);

        // Set the shape's fill color (optional, just for visibility).
        shape.FillFormat.ForeColor = System.Drawing.Color.LightBlue;

        // Embed a tooltip comment by setting the AlternativeText property.
        // This text appears when the mouse hovers over the shape in Excel.
        shape.AlternativeText = "This is a tooltip displayed on hover.";

        // (Optional) Add a cell comment to demonstrate a regular comment.
        // The comment will appear as a note when the cell is selected.
        Comment comment = sheet.Comments[sheet.Comments.Add("B5")] as Comment;
        comment.Note = "Regular cell comment.";

        // Save the workbook to a file.
        workbook.Save("ShapeWithTooltip.xlsx");
    }
}
