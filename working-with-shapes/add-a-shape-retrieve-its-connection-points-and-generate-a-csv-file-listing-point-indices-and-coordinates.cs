// Title: Create a rectangle shape in an Excel worksheet, extract its connection points, and export them to a CSV file using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that adds a rectangle shape to a worksheet with Aspose.Cells, calls GetConnectionPoints, and writes each point's index, X and Y values to a CSV file. | Write a reusable C# method that accepts any Aspose.Cells Shape object, retrieves its connection points, and returns the data as a CSV‑formatted string. | Adapt the example to output the shape's connection points as a JSON array instead of a CSV file using Aspose.Cells.
// Common Searches: how to use Aspose.Cells GetConnectionPoints to list shape coordinates in C# | export rectangle shape connection points to CSV with Aspose.Cells .NET | C# Aspose.Cells example for retrieving shape connection point indices and positions
// Tags: add rectangle shape Aspose.Cells C# | retrieve shape connection points Aspose.Cells | export shape coordinates to CSV C# | Aspose.Cells GetConnectionPoints example | write connection point data Aspose.Cells

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, inserts a rectangle shape on the first worksheet, obtains its connection points via GetConnectionPoints, writes each point's index and X/Y coordinates to a CSV file, and saves the workbook.
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

            // Retrieve the collection of connection points for the shape
            // Use dynamic to avoid compile‑time dependency on ConnectionPointInfo type
            var points = shape.GetConnectionPoints();

            // Path for the CSV file that will contain the connection point data
            string csvPath = "ShapeConnectionPoints.csv";

            // Ensure the directory for the CSV file exists
            string csvDir = Path.GetDirectoryName(csvPath);
            if (!string.IsNullOrEmpty(csvDir) && !Directory.Exists(csvDir))
            {
                Directory.CreateDirectory(csvDir);
            }

            // Write the connection point data to the CSV file
            using (StreamWriter writer = new StreamWriter(csvPath))
            {
                // Write CSV header
                writer.WriteLine("Index,X,Y");

                // Write each connection point's index and coordinates
                foreach (dynamic pt in points)
                {
                    writer.WriteLine($"{pt.Index},{pt.X},{pt.Y}");
                }
            }

            // Save the workbook (optional, just to keep the shape in a file)
            string workbookPath = "ShapeWorkbook.xlsx";

            // Ensure the directory for the workbook exists
            string wbDir = Path.GetDirectoryName(workbookPath);
            if (!string.IsNullOrEmpty(wbDir) && !Directory.Exists(wbDir))
            {
                Directory.CreateDirectory(wbDir);
            }

            workbook.Save(workbookPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
