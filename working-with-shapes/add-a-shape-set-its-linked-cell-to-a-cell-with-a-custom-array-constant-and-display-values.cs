// Title: Create a rectangle shape linked to a cell with a custom array constant and display its values using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add rectangle shape and link it to cell A1 containing an array constant | how to set shape.LinkedCell to a cell with an array formula in Aspose.Cells .NET | display array constant values inside a shape's text with Aspose.Cells for .NET | programmatically add a shape and bind it to a worksheet cell using Aspose.Cells C# | save workbook after linking shape to cell that has ={10,20,30,40} formula Aspose.Cells
// Tags: add rectangle shape Aspose.Cells | link shape to worksheet cell Aspose.Cells | set array constant formula in cell | populate shape text with array values | save workbook as Xlsx Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsDemo
{
    // The example creates a new workbook, writes an array constant formula to cell A1, adds a rectangle shape at row 5 column 5, links the shape to cell A1, sets the shape's text to show the array values, and saves the file as ArrayShapeDemo.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook(); // {CreateWorkbook}
                var worksheet = workbook.Worksheets[0];

                // Set a cell with a custom array constant
                var arrayCell = worksheet.Cells["A1"];
                arrayCell.Formula = "={10,20,30,40}"; // Excel array constant

                // Add a rectangle shape to the worksheet
                // Parameters: shape type, upper-left row, upper-left column,
                // upper-left row offset (pixels), upper-left column offset (pixels),
                // height (pixels), width (pixels)
                Shape shape = worksheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle,
                    5,   // upper-left row
                    5,   // upper-left column
                    0,   // row offset
                    0,   // column offset
                    50,  // height
                    200  // width
                );

                // Link the shape to the cell containing the array constant
                shape.LinkedCell = "A1";

                // Optionally display the array values in the shape's text
                shape.Text = "Array values: 10, 20, 30, 40";

                // Save the workbook
                workbook.Save("ArrayShapeDemo.xlsx", SaveFormat.Xlsx); // {SaveWorkbook}
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
