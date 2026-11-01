// Title: Add a rectangle shape linked to a multi‑line cell and enable text wrapping with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to insert a rectangle shape, bind its text to cell A1 containing line‑break characters, and turn on text wrapping. | Demonstrate how to set a shape's Text property to "={A1}" so the shape displays the multi‑line content of a worksheet cell with proper alignment. | Provide a complete example that creates a workbook, writes a multi‑line string to a cell, adds a linked rectangle shape, configures wrapping, and saves the file.
// Common Searches: aspnet c# aspose.cells bind shape text to cell with line breaks | how to display multi line cell value inside a shape using aspose.cells | aspose.cells shape text wrapping rectangle shape linked to cell | c# create shape that shows cell A1 content with newline characters | aspose.cells set shape.Text = "={A1}" and enable wrapping
// Tags: add rectangle shape linked to cell Aspose.Cells | bind shape text to cell formula Aspose.Cells | enable text wrapping in shape Aspose.Cells | multi‑line cell value display in shape C# | Aspose.Cells shape text alignment configuration

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, writes a multi‑line string to cell A1, adds a rectangle shape, links the shape's text to the cell using the ={A1} formula, enables text wrapping and alignment, and saves the workbook as MultiLineShape.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a multi-line text into cell A1 (use \n for line breaks)
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Line 1\nLine 2\nLine 3");

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column,
            // row offset, column offset, height (pixels), width (pixels)
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,                        // upper left row (zero‑based)
                2,                        // upper left column (zero‑based)
                0,                        // row offset
                0,                        // column offset
                100,                      // height in pixels
                200);                     // width in pixels

            // Link the shape's text to cell A1.
            // The formula syntax "={A1}" makes the shape display the cell's value,
            // including any line break characters.
            shape.Text = "={A1}";

            // Enable text wrapping inside the shape to ensure proper display
            shape.TextHorizontalAlignment = TextAlignmentType.General;
            shape.TextVerticalAlignment = TextAlignmentType.Top;
            shape.IsTextWrapped = true;

            // Save the workbook to a file
            string outputPath = "MultiLineShape.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
