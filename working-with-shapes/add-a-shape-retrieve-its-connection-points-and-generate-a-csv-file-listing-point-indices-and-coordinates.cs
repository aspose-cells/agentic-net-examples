// Title: Add a rectangle shape to an Excel worksheet, read its connection points, and export them to a CSV file with Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to insert a rectangle shape on the first worksheet, calls GetConnectionPoints(), and writes each point's index and X/Y values into a CSV file. | Generate a .NET example that creates an Excel workbook, adds a rectangle shape, enumerates its connection points, and saves the coordinates as a CSV document.
// Common Searches: Aspose.Cells C# get connection points of a shape in Excel | export shape connection point coordinates to CSV using Aspose.Cells .NET | sample code for retrieving rectangle shape connection points with Aspose.Cells | how to list shape connection points in an Excel file via Aspose.Cells C# | save shape connection points as CSV file in .NET Excel library
// Tags: add rectangle shape Aspose.Cells | retrieve shape connection points .NET | export shape coordinates to CSV C# | Aspose.Cells GetConnectionPoints example | write shape points CSV Aspose

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program creates a new workbook, adds a rectangle shape to the first worksheet, obtains its connection points via GetConnectionPoints(), builds a CSV containing index, X and Y values, writes the CSV file, and optionally saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: drawing type, upper left row, upper left column,
            // row offset, column offset, height, width
            Shape shape = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,   // upper left row
                2,   // upper left column
                0,   // row offset (in points)
                0,   // column offset (in points)
                100, // height (in points)
                200  // width (in points)
            );

            // Retrieve the connection points for the shape.
            // Use var to avoid compile‑time dependency on the ConnectionPoint type.
            var connectionPoints = shape.GetConnectionPoints();

            // Build CSV content with header
            StringBuilder csvBuilder = new StringBuilder();
            csvBuilder.AppendLine("Index,X,Y");

            // Iterate through each connection point and write its data
            foreach (var point in connectionPoints)
            {
                // Use dynamic to access properties at runtime
                dynamic p = point;
                csvBuilder.AppendLine($"{p.Index},{p.X},{p.Y}");
            }

            // Define CSV file path
            string csvFilePath = "ShapeConnectionPoints.csv";

            // Write CSV data to file
            File.WriteAllText(csvFilePath, csvBuilder.ToString());

            // Save the workbook (optional, to verify the shape exists)
            string workbookPath = "ShapeWorkbook.xlsx";
            workbook.Save(workbookPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
