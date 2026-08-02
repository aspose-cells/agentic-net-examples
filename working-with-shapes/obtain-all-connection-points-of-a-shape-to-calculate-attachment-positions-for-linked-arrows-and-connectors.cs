// Title: C# – Retrieve Shape Connection Points with Aspose.Cells for Precise Connector Placement
// Description: This example creates a workbook, adds a rectangle shape, calls Shape.GetConnectionPoints() to obtain the X‑Y coordinates of every anchor on the shape, prints them, and saves the file. Use the returned points to align arrows, lines, or other connectors accurately.
// Keywords: Aspose.Cells C# shape connection points | GetConnectionPoints method .NET | worksheet shape anchor coordinates | connector attachment positions Aspose | calculate arrow endpoints programmatically | flowchart diagram automation | Aspose.Cells drawing API
// Common Searches: how to obtain shape anchors in Aspose.Cells | Aspose.Cells GetConnectionPoints usage example | C# retrieve X Y of shape connection points | determine connector start points for Excel shapes | Aspose.Cells diagramming guide
// Developer Intent: Extract the X and Y values of every connection anchor on a worksheet shape.
// Use Cases: Place arrows or lines exactly on a rectangle’s predefined points. | Synchronize multiple shapes by matching their anchor locations. | Update connector endpoints automatically after resizing a shape.
// AI Prompts: Generate C# code that stores each point from shape.GetConnectionPoints() into a List<PointF> for later calculations. | Show how to draw a connector line between two shapes using the retrieved anchor coordinates. | Explain how to refresh connection‑point data after modifying a shape’s dimensions to keep links accurate.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This example creates a workbook, adds a rectangle shape, calls Shape.GetConnectionPoints() to obtain the X‑Y coordinates of every anchor on the shape, prints them, and saves the file. Use the returned points to align arrows, lines, or other connectors accurately.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upperLeftRow, upperLeftColumn, upperLeftRowOffset, upperLeftColumnOffset, height, width
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 100, 200, 0);

        // Retrieve all connection points of the shape
        float[][] points = shape.GetConnectionPoints();

        // Output the connection points (X,Y) to the console
        Console.WriteLine("Connection Points:");
        for (int i = 0; i < points.Length; i++)
        {
            Console.WriteLine($"Point {i + 1}: X={points[i][0]}, Y={points[i][1]}");
        }

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("ConnectionPointsDemo.xlsx");
    }
}
