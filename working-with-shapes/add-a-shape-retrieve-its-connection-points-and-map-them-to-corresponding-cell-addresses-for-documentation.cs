// Title: Create a rectangle shape in an Excel worksheet, retrieve its connection points, and map them to cell addresses using Aspose.Cells for .NET
// AI Prompts: Add a rectangle shape to a worksheet, call GetConnectionPoints, and print each point's corresponding Excel cell reference. | Generate a workbook, insert a shape, enumerate its connection points, translate row/column indices to cell names, and save the file.
// Common Searches: Aspose.Cells .NET retrieve shape connection points example | C# map shape connection point to Excel cell address | GetConnectionPoints usage with rectangle shape Aspose.Cells | convert shape connection point row column to cell name in Aspose.Cells | document shape coordinates in Excel using Aspose.Cells C#
// Tags: Aspose.Cells GetConnectionPoints API | add rectangle shape Aspose.Cells | shape connection points to Excel cell address | convert row column indices to cell name Aspose.Cells | document shape coordinates .NET

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The sample creates a new workbook, adds a rectangle shape at a specific location, retrieves its connection points via GetConnectionPoints, converts each point's row and column indices to the corresponding Excel cell address, outputs the mappings, and saves the workbook as ShapeConnectionPoints.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, row offset, column offset, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 80);

            // Retrieve the shape's connection points
            var points = shape.GetConnectionPoints();

            // Iterate through each connection point
            foreach (dynamic pt in points)
            {
                // Row and Column are provided by the connection point object
                int row = pt.Row;
                int column = pt.Column;

                // Convert row/column indices to cell address (e.g., "C3")
                string cellAddress = sheet.Cells[row, column].Name;

                // Output the mapping
                Console.WriteLine($"Connection point at Row={row}, Column={column} maps to cell {cellAddress}");
            }

            // Save the workbook
            string outputFile = "ShapeConnectionPoints.xlsx";
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputFile)}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
