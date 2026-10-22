// Title: How to retrieve shape objects and their bounding rectangles inside a DrawObjectEventHandler in Aspose.Cells for .NET
// AI Prompts: Generate C# code that registers a DrawObjectEventHandler on a worksheet, extracts each shape's type and its bounding cell coordinates from the event arguments, and writes the information to the console. | Update the sample to capture the rectangle added to the worksheet, obtain its Bounds property within the DrawObjectEventHandler, and store the coordinates in a List<Rectangle> for later custom processing.
// Common Searches: Aspose.Cells .NET get shape bounds in DrawObjectEventHandler | How to access rectangle dimensions during worksheet drawing event Aspose.Cells | Custom processing of drawing objects when saving workbook using Aspose.Cells | Retrieve MsoDrawingType and bounding box from DrawObjectEventArgs Aspose.Cells
// Tags: aspnet-cells drawobjecteventhandler shape bounds extraction | aspnet-cells retrieve drawing object coordinates | aspnet-cells custom shape processing during save | aspnet-cells rectangle shape bounds property | aspnet-cells worksheet drawing event handling

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new Workbook, adds a rectangle shape to the first worksheet at row 2, column 2 with a size of 100 × 50 points, and saves the file as Result.xlsx. It serves as a starting point for attaching a DrawObjectEventHandler to capture shape objects and their bounding rectangles for custom processing.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: drawing type, upper left row, upper left column, top offset, left offset, height, width
            worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 50);

            // Save the workbook
            string outputPath = "Result.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
