// Title: C# – Add a Rectangle Shape with Aspose.Cells, Extract Connection Points, and Export to CSV
// Description: This example creates a new workbook, inserts a rectangle shape on the first worksheet, uses GetConnectionPoints to obtain the shape's X/Y anchor points, and writes each point's index and coordinates to a CSV file. The workbook can be saved optionally to retain the shape.
// Keywords: Aspose.Cells GetConnectionPoints | C# shape connection points | export shape coordinates CSV | add rectangle shape Aspose.Cells | Aspose.Cells shape geometry | write shape data to CSV | Aspose.Cells connection points example
// Common Searches: Aspose.Cells GetConnectionPoints C# | how to export shape connection points to CSV | create rectangle shape in Aspose.Cells .NET | retrieve shape connection coordinates Aspose.Cells | save shape geometry as CSV using Aspose.Cells
// Developer Intent: Create a rectangle shape, read its connection points, and write them to a CSV file.
// Use Cases: Generate a CSV catalog of shape anchor points for integration with external diagram tools. | Persist shape geometry for automated layout validation in reporting pipelines. | Log connection point data to troubleshoot worksheet positioning issues. | Provide coordinate data for custom charting or mapping applications.
// AI Prompts: Generate C# code that adds a rectangle shape to a worksheet, extracts its connection points with GetConnectionPoints, and writes the index and X/Y values to a CSV file using Aspose.Cells. | Describe how the coordinate values returned by GetConnectionPoints map to worksheet rows, columns, and offsets. | Show how to include additional shape attributes such as width, height, and placement type in the CSV output. | Explain how to modify the example to process multiple shapes and aggregate their connection points into a single CSV.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This example creates a new workbook, inserts a rectangle shape on the first worksheet, uses GetConnectionPoints to obtain the shape's X/Y anchor points, and writes each point's index and coordinates to a CSV file. The workbook can be saved optionally to retain the shape.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, top offset, upper left column, left offset, height, width, placement type (0 = MoveAndSize)
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 100, 200, 0);

        // Retrieve the connection points of the shape
        // Each point is a float[2] where [0] = X and [1] = Y
        float[][] connectionPoints = shape.GetConnectionPoints();

        // Write the connection points to a CSV file
        string csvPath = "ConnectionPoints.csv";
        using (StreamWriter writer = new StreamWriter(csvPath))
        {
            // Header
            writer.WriteLine("Index,X,Y");

            // Data rows
            for (int i = 0; i < connectionPoints.Length; i++)
            {
                float x = connectionPoints[i][0];
                float y = connectionPoints[i][1];
                writer.WriteLine($"{i + 1},{x},{y}");
            }
        }

        // Save the workbook (optional, just to keep the shape in the file)
        workbook.Save("ShapeWithConnectionPoints.xlsx");
    }
}
