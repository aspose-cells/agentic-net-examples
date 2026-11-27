// Title: Create a rectangle shape, retrieve its connection points, and map them to Excel cell addresses with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells get shape connection points and convert to A1 cell reference | How to list connection points of a rectangle shape in an Excel worksheet using Aspose.Cells | Map shape connection points to cell addresses in Aspose.Cells .NET example | Retrieve and display shape connection point coordinates with Aspose.Cells for .NET | Add worksheet comment documenting shape connection points using Aspose.Cells
// Tags: add rectangle shape Aspose.Cells | retrieve shape connection points Aspose.Cells | map connection points to cell addresses | document shape connections with worksheet comments | Aspose.Cells shape API .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, adds a rectangle shape to the first worksheet, uses GetConnectionPoints to obtain each connection point, converts the row and column indices of each point to A1 cell notation, writes the mappings to the console, adds a comment to document the connections, and saves the workbook while handling possible exceptions.
class ShapeConnectionPointsDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, upper left offset (pixels),
            // lower right row, lower right column, lower right offset (pixels)
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 120, 60);
            shape.Text = "Demo Shape";

            // Retrieve connection points using dynamic to avoid compile‑time dependency on the exact type
            try
            {
                dynamic connPoints = shape.GetConnectionPoints();

                // Iterate through each connection point and map it to the corresponding cell address
                for (int i = 0; i < connPoints.Count; i++)
                {
                    dynamic cp = connPoints[i];
                    int row = cp.Row;
                    int column = cp.Column;

                    // Get the cell address in A1 notation
                    string cellAddress = sheet.Cells[row, column].Name;

                    // Output the mapping to the console
                    Console.WriteLine($"Connection Point {i}: Row={row}, Column={column} => Cell {cellAddress}");
                }
            }
            catch (Exception ex)
            {
                // If the API is not supported in the current version, inform the user but continue execution
                Console.WriteLine($"Connection points could not be retrieved: {ex.Message}");
            }

            // Optionally, add a comment to the first cell documenting the connection points
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "Shape connection points mapped to cells. See console output for details.";

            // Save the workbook
            string outputFile = "ShapeConnectionPointsDemo.xlsx";
            try
            {
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved to {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
