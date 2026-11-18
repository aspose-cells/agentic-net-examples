// Title: Add a rectangle shape with multi‑line text linked to a worksheet cell containing line breaks using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set shape text from a cell that includes newline characters | How to display multi‑line cell value inside a rectangle shape in an Excel file with Aspose.Cells | Link a shape's text to a worksheet cell with line breaks using Aspose.Cells for .NET | Create a rectangle shape that shows cell B2 multi‑line content in Aspose.Cells C# | Aspose.Cells shape text alignment for multi‑line cell values
// Tags: add rectangle shape Aspose.Cells C# | shape text from cell with line breaks Aspose.Cells | multi‑line text in Excel shape Aspose.Cells | set shape text alignment Aspose.Cells | bind shape to worksheet cell value Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, writes a string with '\n' line breaks into cell B2, adds a rectangle shape, assigns the cell's string value to the shape's Text property, centers the text horizontally and vertically, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put multi‑line text into a cell using line‑break character
            Cell cell = sheet.Cells["B2"];
            cell.PutValue("First line\nSecond line\nThird line");

            // Add a rectangle shape to the worksheet
            // Parameters: type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 5, 0, 5, 0, 200, 100);

            // Link the shape's text to the cell's value (including line breaks)
            shape.Text = cell.StringValue;

            // Ensure the shape displays the multi‑line text correctly
            shape.TextHorizontalAlignment = TextAlignmentType.Center;
            shape.TextVerticalAlignment = TextAlignmentType.Center;

            // Define output file path
            string outputPath = "output.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
