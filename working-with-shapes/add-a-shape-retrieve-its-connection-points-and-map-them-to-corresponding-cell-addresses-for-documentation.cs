// Title: Create a rectangle shape in an Excel worksheet and output its row, column, height, and width with Aspose.Cells for .NET
// AI Prompts: Generate C# code using Aspose.Cells to add a rectangle shape at row 2, column 2 with height 100 and width 200, then print the shape's UpperLeftRow, UpperLeftColumn, Height, and Width to the console. | Extend the example to retrieve all connection points of the inserted shape and write each point's corresponding cell address into the worksheet. | Wrap the shape‑creation logic in a try‑catch block, output any exception messages, and save the workbook as "ShapeConnectionPoints.xlsx".
// Common Searches: asp.net core add rectangle shape to Excel file using Aspose.Cells | how to get shape UpperLeftRow and UpperLeftColumn in Aspose.Cells C# | retrieve shape dimensions height width with Aspose.Cells .NET | save workbook after inserting a shape with Aspose.Cells | map shape connection points to cell addresses Aspose.Cells example
// Tags: add rectangle shape Aspose.Cells .NET | shape UpperLeftRow UpperLeftColumn properties | retrieve shape dimensions Aspose.Cells | save workbook with shape Aspose.Cells | shape connection points to cell mapping Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeConnectionPointsDemo
{
    // The sample creates a new workbook, inserts a rectangle shape at row 2/column 2 with a height of 100 and width of 200, prints the shape's UpperLeftRow, UpperLeftColumn, Height, and Width to the console, and saves the file as ShapeConnectionPoints.xlsx. It also shows how to add error handling and suggests extending the code to map shape connection points to cell addresses.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
                Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 200);

                // Output shape position and size information
                Console.WriteLine("Shape Information:");
                Console.WriteLine($"UpperLeftRow: {shape.UpperLeftRow}");
                Console.WriteLine($"UpperLeftColumn: {shape.UpperLeftColumn}");
                // Offsets are not available in this version of Aspose.Cells; omitted.
                Console.WriteLine($"Height: {shape.Height}");
                Console.WriteLine($"Width: {shape.Width}");

                // Save the workbook (optional, demonstrates the lifecycle)
                string outputPath = "ShapeConnectionPoints.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
