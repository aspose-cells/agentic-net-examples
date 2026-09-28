// Title: Add a rectangle shape in Aspose.Cells for .NET and bind its multiline text to a worksheet cell containing line‑break characters
// AI Prompts: Write C# code that creates a rectangle shape on a worksheet, reads a cell value with '\n' line breaks, assigns it to the shape’s Text property, and configures the shape to display each line separately. | Show how to connect a shape’s displayed text to a worksheet cell that contains newline characters while keeping the line breaks intact when the workbook is saved with Aspose.Cells. | Demonstrate setting horizontal and vertical alignment for a shape that shows multiline text linked to a cell in Aspose.Cells.
// Common Searches: aspnet how to bind shape text to a cell with newline characters using Aspose.Cells | c# Aspose.Cells rectangle shape multiline text from cell value | preserve line breaks in shape text Aspose.Cells for .NET | display cell text with \n inside a shape in Excel using Aspose.Cells | set shape text alignment for multiline content Aspose.Cells C#
// Tags: Aspose.Cells add rectangle shape | bind shape text to cell Aspose.Cells | shape text with line breaks Aspose.Cells | shape text alignment Aspose.Cells | C# Aspose.Cells shape linking

using Aspose.Cells;
using Aspose.Cells.Drawing;

// // This program creates a workbook, writes a cell with newline characters, adds a rectangle shape, assigns the cell's string value to the shape's Text property, centers the text horizontally and vertically, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];

        // Insert multi-line text into a cell using line break character (char 10)
        sheet.Cells["A1"].PutValue("First line\nSecond line\nThird line");

        // Add a rectangle shape to the worksheet
        // Parameters: type, upper left row, upper left column, top offset, left offset, height, width
        var shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 0, 0, 0, 100, 200);

        // Link the shape's text to the cell's value
        shape.Text = sheet.Cells["A1"].StringValue;

        // Ensure the shape displays the multi-line text correctly
        shape.TextHorizontalAlignment = TextAlignmentType.Center;
        shape.TextVerticalAlignment = TextAlignmentType.Center;

        // Save the workbook
        workbook.Save("output.xlsx");
    }
}
