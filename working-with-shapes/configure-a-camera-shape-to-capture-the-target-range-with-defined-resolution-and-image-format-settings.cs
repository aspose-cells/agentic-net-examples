// Title: Insert a rectangle shape as a camera placeholder linked to range A1:C5 and save as XLSX using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a new workbook, fills cells A1:C5 with sample text, adds a rectangle shape at D1 sized 400 × 300 points, names it "CameraShapePlaceholder", and saves the file as CameraShapeExample.xlsx using Aspose.Cells. | Write a C# snippet that defines a cell range, inserts a rectangle shape to represent a camera, sets its dimensions in points, links the shape to the defined range, and exports the workbook to XLSX with Aspose.Cells. | Provide C# instructions to add a placeholder shape for a camera to an Aspose.Cells worksheet, associate it with a specific range, and handle the absence of a native CameraShape class when saving the workbook.
// Common Searches: how to add a rectangle shape as a camera placeholder in Aspose.Cells C# | Aspose.Cells set shape size in points and associate with cell range | C# Aspose.Cells create shape at specific cell and save workbook | Aspose.Cells example inserting shape linked to range A1:C5 | saving workbook with custom shape using Aspose.Cells for .NET
// Tags: shape insertion Aspose.Cells C# | set shape dimensions points Aspose.Cells | associate shape with cell range Aspose.Cells | camera representation shape Aspose.Cells | export workbook with custom shape Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;

// The program creates a new workbook, populates cells A1:C5 with sample values, defines that range, adds a rectangle shape at D1 sized 400 × 300 points, names it "CameraShapePlaceholder" (noting that Aspose.Cells lacks a native CameraShape class), and saves the file as CameraShapeExample.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // Fill sample data in the range A1:C5
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Define the range that would be captured by a camera (if supported)
            var targetRange = sheet.Cells.CreateRange("A1:C5");

            // Add a rectangle shape at cell D1 (row 0, column 3) with size 400x300 points.
            // Parameters: type, upperLeftRow, upperLeftColumn, top, left, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                0,                        // upper left row (D1 -> row 0)
                3,                        // upper left column (D -> column 3)
                0,                        // top offset in points
                0,                        // left offset in points
                300,                      // height in points
                400);                     // width in points

            shape.Name = "CameraShapePlaceholder";

            // Note: Aspose.Cells for .NET does not expose a CameraShape class in current versions.
            // The placeholder shape can be used to represent a camera-like object.

            // Save the workbook
            string outputPath = "CameraShapeExample.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
