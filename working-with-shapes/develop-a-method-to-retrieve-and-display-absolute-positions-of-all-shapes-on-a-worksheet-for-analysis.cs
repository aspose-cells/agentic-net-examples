// Title: Retrieve absolute positions of all shapes on each worksheet with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to list shape left and top coordinates for every worksheet | Get width and height of Excel shapes using Aspose.Cells API | Iterate through shapes collection and output their positions in points with Aspose.Cells | Display shape geometry (left, top, width, height) from an Excel file in a .NET console app | Extract absolute position values of all drawing objects in a workbook using Aspose.Cells
// Tags: Aspose.Cells shape coordinate extraction | C# iterate worksheet shapes collection | retrieve shape dimensions in points | Excel drawing objects position retrieval | console output of shape geometry Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an Excel workbook, loops through each worksheet and its Shapes collection, and writes each shape's Name together with its Left, Top, Width, and Height values (measured in points) to the console.
class ShapePositionRetriever
{
    // Retrieves and displays absolute positions of all shapes in each worksheet of the workbook.
    public static void DisplayShapePositions(string filePath)
    {
        // Load the workbook from the specified file.
        Workbook workbook = new Workbook(filePath);

        // Iterate through all worksheets.
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Console.WriteLine($"Worksheet: {sheet.Name}");

            // Iterate through all shapes on the current worksheet.
            foreach (Shape shape in sheet.Shapes)
            {
                // Absolute position values are in points (1 point = 1/72 inch).
                double left   = shape.Left;   // Distance from the left edge of the worksheet.
                double top    = shape.Top;    // Distance from the top edge of the worksheet.
                double width  = shape.Width;  // Width of the shape.
                double height = shape.Height; // Height of the shape.

                Console.WriteLine($"Shape Name: {shape.Name}");
                Console.WriteLine($"  Left: {left} pt, Top: {top} pt, Width: {width} pt, Height: {height} pt");
            }
        }
    }

    // Example entry point.
    static void Main(string[] args)
    {
        if (args.Length == 0)
        {
            Console.WriteLine("Please provide the path to the Excel file as a command‑line argument.");
            return;
        }

        string filePath = args[0];
        DisplayShapePositions(filePath);
    }
}
