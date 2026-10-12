// Title: Add a rectangle shape to an Excel worksheet, set its top‑left corner using pixel coordinates, and verify placement with Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a rectangle shape on the first worksheet, converts given pixel X/Y values (e.g., 150, 200) to points with the 0.75 factor, assigns the Left and Top properties, and checks that the resulting coordinates match the expected points. | Write a C# snippet using Aspose.Cells to add a shape, define its width and height in points, position it based on pixel offsets, validate the placement tolerance, and save the workbook as an .xlsx file.
// Common Searches: asp.net how to place a shape at exact pixel location in Excel using Aspose.Cells | convert pixel coordinates to points for shape positioning in Aspose.Cells C# | verify rectangle shape coordinates after setting Left and Top in Aspose.Cells workbook | set shape size in points and position using pixel offsets Aspose.Cells .NET example
// Tags: add rectangle shape Aspose.Cells C# | shape left top pixel conversion Aspose.Cells | pixel to point conversion Excel shape positioning | validate shape placement Aspose.Cells | save workbook as xlsx Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program creates a new workbook, adds a rectangle shape to the first worksheet, sets its width and height in points, converts pixel coordinates (150, 200) to points using a 0.75 factor, assigns the Left and Top properties, validates the placement accuracy, outputs the result, and saves the file as ShapePlacement.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape directly (returns the Shape object)
            Shape shape = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                0,   // upper left row
                0,   // upper left column
                0,   // row offset (in pixels)
                0,   // column offset (in pixels)
                0,   // height (in points) – will be set later
                0);  // width (in points) – will be set later

            // Set shape size (points)
            shape.Width = 100;   // width in points
            shape.Height = 50;   // height in points

            // Desired pixel coordinates for the top‑left corner
            int pixelX = 150; // horizontal pixel position
            int pixelY = 200; // vertical pixel position

            // Convert pixels to points (1 pixel ≈ 0.75 point at 96 DPI)
            const double pixelsToPoints = 0.75;
            shape.Left = (int)(pixelX * pixelsToPoints);
            shape.Top = (int)(pixelY * pixelsToPoints);

            // Validate placement
            bool isPlacementCorrect = Math.Abs(shape.Left - pixelX * pixelsToPoints) < 0.01 &&
                                      Math.Abs(shape.Top - pixelY * pixelsToPoints) < 0.01;

            Console.WriteLine(isPlacementCorrect
                ? "Shape placed correctly."
                : "Shape placement error.");

            // Save the workbook
            workbook.Save("ShapePlacement.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
