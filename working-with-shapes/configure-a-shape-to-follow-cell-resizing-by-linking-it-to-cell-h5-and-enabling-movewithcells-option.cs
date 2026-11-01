// Title: How to make a rectangle shape move and resize with cell H5 using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# example to bind a shape to a specific worksheet cell | set shape placement to MoveAndSize so it follows cell resizing in Aspose.Cells | link a rectangle shape to cell H5 and keep it aligned during row/column changes | using PlacementType.MoveAndSize with Aspose.Cells to attach shapes to cells
// Tags: Aspose.Cells shape placement MoveAndSize | C# anchor rectangle shape to worksheet cell | shape follows cell resizing Aspose.Cells | PlacementType usage in Aspose.Cells .NET | add shape to cell H5 example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, adds a rectangle shape anchored to cell H5 on the first worksheet, sets its Placement property to MoveAndSize so the shape moves and resizes together with the cell, and saves the workbook as output.xlsx.
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

            // Add a rectangle shape anchored to cell H5 (row index 4, column index 7)
            // Use MsoDrawingType for compatibility with older Aspose.Cells versions
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 4, 7, 0, 0, 100, 30);

            // Ensure the shape moves and resizes with the cell
            shape.Placement = PlacementType.MoveAndSize;

            // Prepare output path
            string outputPath = "output.xlsx";

            // Determine directory (use current directory if none specified)
            string outputDir = Path.GetDirectoryName(outputPath) ?? Directory.GetCurrentDirectory();

            // Create directory if it does not exist
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
