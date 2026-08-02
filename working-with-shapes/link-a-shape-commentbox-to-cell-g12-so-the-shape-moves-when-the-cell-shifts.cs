// Title: Link a TextBox shape to cell G12 so it moves with the cell – Aspose.Cells for .NET
// Description: Shows how to create a workbook, add a TextBox at G6, set its text, link it to the absolute cell $G$12 using the LinkedCell property, and save the file, ensuring the shape stays attached when rows or columns shift.
// Keywords: Aspose.Cells | LinkedCell property | shape linking | textbox comment box | C# | .NET | move shape with cell | worksheet layout changes | dynamic annotation
// Common Searches: Aspose.Cells link shape to cell | C# link textbox to cell G12 | move shape with cell in Aspose.Cells | LinkedCell usage example | shape follows cell after insert rows
// Developer Intent: Link a textbox shape to cell G12 so the shape moves when the cell shifts.
// Use Cases: Attach a comment box to a cell so it remains aligned after inserting or deleting rows/columns. | Create annotations that travel with data when the worksheet is sorted or filtered. | Maintain consistent layout in reporting templates by positioning shapes relative to specific cells.
// AI Prompts: Generate C# code with Aspose.Cells that adds a rectangle shape and links it to cell B5 so it moves with the cell. | Show how to change an existing shape's LinkedCell from $G$12 to $H$15 using Aspose.Cells for .NET. | Explain the impact of the LinkedCell property on shape positioning during row or column operations in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to create a workbook, add a TextBox at G6, set its text, link it to the absolute cell $G$12 using the LinkedCell property, and save the file, ensuring the shape stays attached when rows or columns shift.
class LinkShapeToCell
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a textbox shape that will act as a comment box
        // Parameters: upper left row, upper left column, top offset, left offset, width, height
        TextBox commentBox = sheet.Shapes.AddTextBox(5, 6, 0, 0, 120, 60); // Row 5 (zero‑based), Column 6 => G6 as starting position

        // Set the text of the comment box (optional)
        commentBox.Text = "This is a comment.";

        // Link the shape to cell G12 so it moves with the cell
        // Use absolute A1 style reference
        commentBox.LinkedCell = "$G$12";

        // Save the workbook
        workbook.Save("LinkedCommentBox.xlsx");
    }
}
