// Title: Link a rectangle shape to a TODAY() cell and format the displayed text with a custom date pattern using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# link shape text to a cell formula | set rectangle shape text to display a date from a worksheet cell | apply a custom date format to shape text linked to a cell in Aspose.Cells | bind a shape to a TODAY() formula cell and style the text | create an Excel shape that shows a formatted date from a cell using Aspose.Cells | Aspose.Cells example linking a shape to a cell and customizing its font
// Tags: shape cell reference linking Aspose.Cells .NET | rectangle shape text binding Aspose.Cells | custom date pattern formatting Aspose.Cells | shape font configuration Aspose.Cells | export workbook with linked shape Aspose.Cells | TODAY() cell formatting Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example creates a new workbook, sets cell A1 to the TODAY() formula with a custom "dd-MMM-yyyy" format, adds a rectangle shape, links the shape's displayed text to cell A1 using "=A1", applies font styling to the shape, and saves the file as LinkedShapeWithDate.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Set a date formula in cell A1 (e.g., today's date)
                Cell dateCell = worksheet.Cells["A1"];
                dateCell.Formula = "TODAY()";

                // Apply a custom date format to the cell (e.g., "dd-MMM-yyyy")
                Style dateStyle = workbook.CreateStyle();
                dateStyle.Custom = "dd-MMM-yyyy";
                dateCell.SetStyle(dateStyle);

                // Add a rectangle shape to the worksheet
                // Parameters: shape type, upper left row, upper left column, top, left, height, width
                Shape shape = worksheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle, // shape type
                    2,    // upper left row
                    0,    // upper left column
                    5,    // top offset (pixels)
                    5,    // left offset (pixels)
                    100,  // height (pixels)
                    200   // width (pixels)
                );

                // Link the shape's displayed text to the cell with the date formula
                shape.Text = "=A1";

                // Set the shape's text formatting (font, size, etc.)
                shape.Font.Name = "Calibri";
                shape.Font.Size = 12;

                // Determine output file path
                string outputPath = "LinkedShapeWithDate.xlsx";

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
}
