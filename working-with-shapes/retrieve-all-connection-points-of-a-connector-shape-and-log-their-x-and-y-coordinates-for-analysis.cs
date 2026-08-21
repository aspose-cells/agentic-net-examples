// Title: C# – Retrieve Connector Shape Connection Points (X,Y) with Aspose.Cells for .NET
// Description: Creates a workbook, adds a line shape that acts as a connector, calls GetConnectionPoints to obtain all X‑Y coordinate pairs, logs each point to the console, and saves the file. Demonstrates how to extract connector geometry for analysis.
// Keywords: Aspose.Cells GetConnectionPoints | C# connector shape coordinates | Aspose.Cells line shape X Y | Excel shape connection points .NET | retrieve connector geometry Aspose | log connector points C# | Aspose.Cells shape API
// Common Searches: Aspose.Cells get connection points of a line shape | C# retrieve connector coordinates in Excel | How to use GetConnectionPoints with Aspose.Cells | Log X Y values of connector shape Aspose | Extract connector geometry from workbook
// Developer Intent: Extract every connection point of a connector (line) shape and output its X and Y coordinates using Aspose.Cells for .NET.
// Use Cases: Validate diagram layout by comparing connector points to expected cell positions. | Export connector geometry for automated layout audits or reporting. | Perform collision detection or alignment checks in complex Excel drawings.
// AI Prompts: Generate a C# method that returns a List<PointF> of all connection points for any Aspose.Cells shape. | Create code that writes connector X‑Y coordinates to a CSV file with error handling for non‑connector shapes. | Show how to overlay markers on a worksheet at each connection point returned by GetConnectionPoints.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, adds a line shape that acts as a connector, calls GetConnectionPoints to obtain all X‑Y coordinate pairs, logs each point to the console, and saves the file. Demonstrates how to extract connector geometry for analysis.
class RetrieveConnectorPoints
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a line shape (acts as a connector) to the worksheet.
            // The line is defined by the start cell (C3) and end cell (F6) – zero‑based indices.
            // Width and height parameters are required; set to 0 for a simple connector.
            Shape connector = worksheet.Shapes.AddLine(2, 2, 5, 5, 0, 0);

            // Retrieve all connection points of the line shape
            float[][] connectionPoints = connector.GetConnectionPoints();

            // Log the X and Y coordinates of each connection point
            Console.WriteLine("Connector Connection Points:");
            for (int i = 0; i < connectionPoints.Length; i++)
            {
                Console.WriteLine($"Point {i + 1}: X = {connectionPoints[i][0]}, Y = {connectionPoints[i][1]}");
            }

            // Save the workbook (optional, just to demonstrate full lifecycle)
            string outputPath = "ConnectorPoints.xlsx";

            // Ensure the directory exists before saving
            string fullPath = Path.GetFullPath(outputPath);
            string outputDir = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {fullPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
