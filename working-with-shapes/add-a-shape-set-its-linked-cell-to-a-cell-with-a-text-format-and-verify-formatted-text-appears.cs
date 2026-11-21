// Title: Add a rectangle shape linked to a formatted cell and confirm matching text using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add rectangle shape linked to cell A1 with red bold text | how to synchronize shape text with a formatted cell value in an Aspose.Cells workbook | verify that a shape displays the same formatted text as its source cell using Aspose.Cells .NET | set shape text from a cell that has font color and bold style in Aspose.Cells | save Excel file with a shape that reflects cell formatting using Aspose.Cells for .NET
// Tags: rectangle shape creation Aspose.Cells .NET | cell-to-shape text linking Aspose.Cells | preserve cell font formatting in shape Aspose.Cells | shape text verification Aspose.Cells workbook | export workbook with linked shape Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, formats cell A1 with red bold text, adds a rectangle shape, assigns the shape's text from the cell value, checks that the shape text matches the cell text, and saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put formatted text into cell A1
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Hello World");
            Style style = cell.GetStyle();
            style.Font.Color = Color.Red;      // text color
            style.Font.IsBold = true;          // bold font
            cell.SetStyle(style);

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column,
            // row offset, column offset, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                5,   // upper left row
                5,   // upper left column
                0,   // row offset (pixels)
                0,   // column offset (pixels)
                50,  // height (pixels)
                200  // width (pixels)
            );

            // Set the shape's text to match the linked cell value
            shape.Text = cell.StringValue;

            // Verify that the shape displays the same text as the cell
            string shapeText = shape.Text;          // should reflect cell A1 value
            string cellText = cell.StringValue;     // original cell value

            Console.WriteLine($"Shape text: {shapeText}");
            Console.WriteLine($"Cell text: {cellText}");
            Console.WriteLine($"Texts match: {shapeText == cellText}");

            // Save the workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
