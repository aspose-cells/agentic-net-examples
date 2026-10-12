// Title: Add a rectangle shape linked to a bold red formatted cell and verify the displayed text with Aspose.Cells for .NET
// AI Prompts: Generate C# code that formats cell A1 with bold red font, inserts a rectangle shape, sets its LinkedCell to A1, and prints the shape's Text property. | Write a C# example using Aspose.Cells to create a workbook, apply styling to a cell, attach a shape to that cell, and output both the shape text and the cell's font attributes for verification.
// Common Searches: aspnet add rectangle shape linked to a cell with bold red text using Aspose.Cells | how to display formatted cell text in a linked shape with Aspose.Cells C# | verify shape text matches linked cell style Aspose.Cells .NET example | Aspose.Cells LinkedCell property does not preserve cell formatting
// Tags: linked rectangle shape Aspose.Cells | bold red font cell Aspose.Cells | verify shape text matches linked cell Aspose.Cells | c# Aspose.Cells shape linking example

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a workbook, formats cell A1 with bold red text, adds a rectangle shape, links the shape to A1, reads the shape's Text property, confirms the cell’s font is bold and red, and saves the file as ShapeLinkedCell.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // 1. Set up a cell with text and apply formatting
            // -------------------------------------------------
            Cell targetCell = sheet.Cells["A1"];
            targetCell.PutValue("Formatted Text");

            // Apply bold and red color to the cell's font
            Style cellStyle = targetCell.GetStyle();
            cellStyle.Font.IsBold = true;
            cellStyle.Font.Color = Color.Red;
            targetCell.SetStyle(cellStyle);

            // -------------------------------------------------
            // 2. Add a shape (rectangle) to the worksheet
            // -------------------------------------------------
            // Parameters: shape type, upper left row, upper left column,
            // upper left row offset, upper left column offset,
            // height (pixels), width (pixels)
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                5, 0,                     // upper left row, column
                5, 0,                     // row offset, column offset
                100, 200);                // height, width

            // -------------------------------------------------
            // 3. Link the shape to the formatted cell
            // -------------------------------------------------
            shape.LinkedCell = "A1";

            // -------------------------------------------------
            // 4. Verify that the shape displays the cell's text
            // -------------------------------------------------
            string displayedText = shape.Text; // Should be "Formatted Text"
            Console.WriteLine("Shape displayed text: " + displayedText);

            // -------------------------------------------------
            // 5. Verify that the linked cell retains its formatting
            // -------------------------------------------------
            Style linkedCellStyle = sheet.Cells[shape.LinkedCell].GetStyle();
            bool isBold = linkedCellStyle.Font.IsBold;
            Color fontColor = linkedCellStyle.Font.Color;

            Console.WriteLine("Linked cell is bold: " + isBold);
            Console.WriteLine("Linked cell font color (ARGB): " + fontColor.ToArgb());

            // -------------------------------------------------
            // 6. Save the workbook (optional)
            // -------------------------------------------------
            workbook.Save("ShapeLinkedCell.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
