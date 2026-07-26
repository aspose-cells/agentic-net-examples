// Title: Retrieve X/Y Coordinates of Connector Shape Connection Points using Aspose.Cells for .NET
// Description: Creates a workbook, adds a line connector shape, calls Shape.GetConnectionPoints() to obtain all connection points as a float array, and logs each point's X and Y values to the console before saving the file.
// Keywords: Aspose.Cells | .NET | C# | connector shape | GetConnectionPoints | Excel line shape | shape connection points | retrieve coordinates | log X Y values | Excel drawing API
// Common Searches: Aspose.Cells get connector shape points | C# retrieve connection points of line shape | Shape.GetConnectionPoints example | log connector coordinates Aspose.Cells | Excel shape connection points .NET
// Developer Intent: Extract all connection points of a connector shape and output their X and Y coordinates.
// Use Cases: Validate diagram geometry by checking exact connector endpoint locations. | Programmatically align or attach other shapes based on connector points. | Export connector coordinates to a report, CSV, or database for further analysis.
// AI Prompts: Write C# code that uses Aspose.Cells to get a connector shape's connection points and stores them in a List<PointF>. | Show how to iterate over Shape.GetConnectionPoints() and write each X,Y pair to a CSV file with Aspose.Cells. | Explain how to modify a connector shape's connection points after retrieving them using GetConnectionPoints.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, adds a line connector shape, calls Shape.GetConnectionPoints() to obtain all connection points as a float array, and logs each point's X and Y values to the console before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a connector (line) shape to the worksheet
            // Parameters: type, upper left row, upper left column, top offset, left offset, height, width
            Shape connector = worksheet.Shapes.AddShape(
                MsoDrawingType.Line,   // connector type
                2,                     // upper left row
                5,                     // upper left column
                5,                     // top offset (in points)
                5,                     // left offset (in points)
                200,                   // height (length of the line)
                0);                    // width (thickness)

            // Retrieve all connection points of the connector shape
            float[][] connectionPoints = connector.GetConnectionPoints();

            // Log the X and Y coordinates of each connection point
            Console.WriteLine("Connector Connection Points:");
            for (int i = 0; i < connectionPoints.Length; i++)
            {
                Console.WriteLine($"Point {i + 1}: X = {connectionPoints[i][0]}, Y = {connectionPoints[i][1]}");
            }

            // Save the workbook to a file
            string outputPath = "ConnectorConnectionPoints.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
