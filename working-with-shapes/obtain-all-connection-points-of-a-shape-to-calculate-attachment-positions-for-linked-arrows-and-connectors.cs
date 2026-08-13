// Title: Aspose.Cells C# – Get All Shape Connection Points for Arrow and Connector Placement
// Description: This example creates a new workbook, adds a rectangle shape, calls Shape.GetConnectionPoints() to retrieve the X‑Y coordinates of every connection point, prints each point to the console, and saves the file as ConnectionPointsDemo.xlsx. Use the returned float[][] to calculate precise attachment locations for arrows, connectors, or custom diagram elements.
// Keywords: Aspose.Cells GetConnectionPoints | C# shape connection points | Excel shape attachment coordinates | Aspose.Cells connector positioning | retrieve shape connection points .NET | diagram arrows Aspose.Cells | Excel drawing API C#
// Common Searches: Aspose.Cells get shape connection points C# | How to obtain connector points from a shape in Aspose.Cells | Shape.GetConnectionPoints example for .NET | Calculate arrow attachment positions in Excel using Aspose.Cells | Retrieve rectangle connection points Aspose.Cells
// Developer Intent: Extract the coordinates of every connection point on a shape to determine where arrows or connectors should attach.
// Use Cases: Loop through the float[][] to draw connector lines between multiple shapes based on the nearest points. | Align SmartArt or flow‑chart elements automatically when generating Excel reports. | Export shape connection data to external diagramming tools for custom rendering.
// AI Prompts: Generate C# code that adds several shapes and draws connectors between the closest connection points using Aspose.Cells. | Show how to find the nearest connection point between two shapes and create a connector line in a workbook. | Explain how resizing a shape affects its connection points and how to recalculate them with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsConnectionPointsDemo
{
    // This example creates a new workbook, adds a rectangle shape, calls Shape.GetConnectionPoints() to retrieve the X‑Y coordinates of every connection point, prints each point to the console, and saves the file as ConnectionPointsDemo.xlsx. Use the returned float[][] to calculate precise attachment locations for arrows, connectors, or custom diagram elements.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, top offset, left offset, height, width, shape type
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 100, 200, 0);

            // Retrieve all connection points of the shape
            float[][] connectionPoints = shape.GetConnectionPoints();

            // Output the connection points – these can be used to attach arrows/connectors
            Console.WriteLine("Connection Points:");
            for (int i = 0; i < connectionPoints.Length; i++)
            {
                float x = connectionPoints[i][0];
                float y = connectionPoints[i][1];
                Console.WriteLine($"Point {i + 1}: X = {x}, Y = {y}");
            }

            // Save the workbook (lifecycle save)
            workbook.Save("ConnectionPointsDemo.xlsx");
        }
    }
}
