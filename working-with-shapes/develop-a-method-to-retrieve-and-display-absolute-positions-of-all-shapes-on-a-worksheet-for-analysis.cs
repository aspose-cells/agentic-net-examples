// Title: C# – Get Absolute Pixel Positions and Sizes of All Shapes in an Aspose.Cells Worksheet
// Description: Loads a workbook, iterates the worksheet’s ShapeCollection, and prints each shape’s name, type, X/Y pixel coordinates and width/height. The example also demonstrates creating a workbook, adding rectangle, oval and line shapes, saving it, and then analyzing the saved file to retrieve shape positions.
// Keywords: Aspose.Cells shape position C# | Aspose.Cells get shape coordinates | Aspose.Cells shape size properties | enumerate worksheet shapes Aspose.Cells | absolute pixel location of shapes Aspose.Cells | .NET Aspose.Cells shape collection | retrieve shape X Y width height Aspose.Cells
// Common Searches: How to read shape X and Y coordinates in Aspose.Cells for .NET | Aspose.Cells C# list all shapes on a worksheet | Get shape width and height using Aspose.Cells | Display shape name, type and position with Aspose.Cells | Aspose.Cells absolute position of shapes in pixels
// Developer Intent: Obtain and display the absolute X/Y pixel coordinates and dimensions of every shape on a worksheet.
// Use Cases: Verify that chart and image placements match design specifications before publishing reports. | Export shape layout data to CSV or JSON for external diagram analysis. | Filter shapes by type (e.g., rectangles) and process only their coordinates in a custom workflow.
// AI Prompts: Generate a method that returns a list of objects containing shape name, type, X, Y, width, and height using Aspose.Cells. | Create a utility that writes each shape’s absolute position and size to a CSV file instead of the console. | Adapt the sample to skip non‑rectangle shapes and output only rectangle coordinates.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads a workbook, iterates the worksheet’s ShapeCollection, and prints each shape’s name, type, X/Y pixel coordinates and width/height. The example also demonstrates creating a workbook, adding rectangle, oval and line shapes, saving it, and then analyzing the saved file to retrieve shape positions.
class ShapePositionAnalyzer
{
    // Loads a workbook and prints absolute positions of all shapes on the first worksheet
    public static void Analyze(string workbookPath)
    {
        // Load the workbook (load rule)
        Workbook workbook = new Workbook(workbookPath);
        Worksheet worksheet = workbook.Worksheets[0];

        ShapeCollection shapes = worksheet.Shapes;

        Console.WriteLine($"Total shapes: {shapes.Count}");
        for (int i = 0; i < shapes.Count; i++)
        {
            Shape shape = shapes[i];

            // Absolute position in pixels (X, Y) and size (Width, Height)
            int x = shape.X;
            int y = shape.Y;
            int width = shape.Width;
            int height = shape.Height;

            Console.WriteLine($"Shape {i}: Name=\"{shape.Name}\", Type={shape.Type}");
            Console.WriteLine($"  Position: X={x} px, Y={y} px");
            Console.WriteLine($"  Size: Width={width} px, Height={height} px");
        }
    }

    // Demonstrates creating a workbook with shapes, saving it, and then analyzing positions
    public static void Main()
    {
        // Create a new workbook and add some shapes (create rule)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle
        worksheet.Shapes.AddRectangle(2, 0, 2, 0, 100, 150);
        // Add an oval
        worksheet.Shapes.AddOval(5, 0, 5, 0, 80, 120);
        // Add a line
        worksheet.Shapes.AddLine(8, 0, 8, 0, 200, 0);

        string filePath = "ShapesDemo.xlsx";

        // Save the workbook (save rule)
        workbook.Save(filePath);

        // Analyze and display shape positions
        Analyze(filePath);
    }
}
