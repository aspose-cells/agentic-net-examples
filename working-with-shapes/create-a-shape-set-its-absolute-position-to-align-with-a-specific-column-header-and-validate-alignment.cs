// Title: Create a rectangle shape and align it to the column B header in an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Insert a rectangle shape into a worksheet and set its UpperLeftRow, UpperLeftColumn, Top, and Left properties so that it sits exactly over the header cell in column B. | Verify that the shape’s position properties match the target cell coordinates and then save the workbook as an .xlsx file.
// Common Searches: Aspose.Cells how to position a shape over a specific column header in C# | set shape absolute coordinates by row and column Aspose.Cells .NET | align rectangle shape to cell B1 using Aspose.Cells API | validate shape alignment with header cell in Excel using Aspose.Cells | C# example for adding and aligning shapes to worksheet cells with Aspose.Cells
// Tags: Aspose.Cells add rectangle shape | Aspose.Cells set shape absolute position | Aspose.Cells align shape to cell | Aspose.Cells validate shape coordinates | Aspose.Cells save workbook as xlsx

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The code creates a new workbook, writes a header in column B, adds a rectangle shape, sets its UpperLeftRow, UpperLeftColumn, Top, and Left properties to align the shape with the header cell, checks the alignment, and saves the file as ShapeAlignment.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Put a header text in column B (index 1) of the first row (index 0)
            worksheet.Cells[0, 1].PutValue("Header B");

            // Add a rectangle shape to the worksheet
            // Parameters: drawing type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                0,   // upper left row
                0,   // upper left column
                0,   // top offset (in points)
                0,   // left offset (in points)
                30,  // height (in points)
                100  // width (in points)
            );

            // Align the shape with the column B header (row 0, column 1)
            shape.UpperLeftRow = 0;      // upper left row
            shape.UpperLeftColumn = 1;   // upper left column
            shape.Top = 0;               // top offset
            shape.Left = 0;              // left offset

            // Validate alignment using shape properties
            bool isAligned = shape.UpperLeftRow == 0 &&
                             shape.UpperLeftColumn == 1 &&
                             shape.Top == 0 &&
                             shape.Left == 0;

            Console.WriteLine("Shape aligned with column B header: " + isAligned);

            // Save the workbook (ensure the directory exists)
            string outputPath = "ShapeAlignment.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine("Error saving workbook: " + saveEx.Message);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
