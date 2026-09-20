// Title: How to link a rectangle shape to a DATEVALUE cell and display the formatted date using Aspose.Cells for .NET (C#)
// AI Prompts: Create a rectangle shape, assign its Text property a formula that points to a DATEVALUE cell, and make the shape display the cell’s formatted date with Aspose.Cells in C#. | Apply a built‑in date style to a worksheet cell, bind a shape’s text to that cell via a formula, and set the shape’s placement to MoveAndSize in a .NET workbook.
// Common Searches: Aspose.Cells C# shape text bound to a DATEVALUE cell | display formatted date inside a rectangle shape using Aspose.Cells | set shape placement to MoveAndSize after linking to a date cell in Aspose.Cells | apply built‑in date format to a cell and reference it from a shape in C#
// Tags: shape text formula reference Aspose.Cells | rectangle shape DATEVALUE linking C# | apply built‑in date format to cell Aspose.Cells | shape placement MoveAndSize Aspose.Cells | Aspose.Cells C# date cell formatting

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, writes a DATEVALUE formula in cell A1, applies a built‑in date format, adds a rectangle shape, links the shape's text to the cell using a formula, configures the shape to move and size with the cell, and saves the file as LinkedShapeDate.xlsx.
class LinkShapeToDateCell
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Put a DATEVALUE formula in cell A1
            Cell dateCell = sheet.Cells["A1"];
            dateCell.Formula = @"=DATEVALUE(""2023-09-19"")";

            // Apply a date format to the cell so the displayed value is formatted
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Number = 14; // Built‑in date format (e.g., mm/dd/yyyy)
            dateCell.SetStyle(dateStyle);

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 0, 0, 0, 60, 200);

            // Link the shape's text to the cell A1 using a formula
            shape.Text = "=A1";

            // Ensure the shape moves and sizes with the cell (optional)
            shape.Placement = PlacementType.MoveAndSize;

            // Save the workbook
            workbook.Save("LinkedShapeDate.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
