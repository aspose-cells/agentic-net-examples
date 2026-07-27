// Title: Get Shape Absolute Position, Convert Pixels to Centimeters, and Write to Cells – Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook with Aspose.Cells for .NET, add a rectangle shape, read its absolute Left and Top pixel offsets, convert those offsets to centimeters (using a 96 DPI reference), write the formatted results into cells A1 and B1, and save the workbook.
// Keywords: Aspose.Cells shape position | pixel to cm conversion C# | shape left top coordinates | write values to Excel cells Aspose.Cells | C# shape absolute location | 96 DPI pixel conversion
// Common Searches: Aspose.Cells get shape left coordinate | convert shape pixel coordinates to centimeters C# | store shape position in Excel cell Aspose | pixel to centimeter conversion Aspose.Cells | shape absolute location .NET
// Developer Intent: The developer needs to obtain a shape's absolute pixel coordinates, translate them into centimeters, and record the results in specific worksheet cells.
// Use Cases: Generate a layout report that lists each shape's position in centimeters for precise printing. | Create a verification sheet that logs shape coordinates to aid alignment checks during automated document generation. | Export shape location data for downstream quality‑control or analytics pipelines.
// AI Prompts: Write C# code with Aspose.Cells that iterates over all shapes in a worksheet, converts their pixel positions to inches, and outputs the values to a summary table. | Provide an example that reads a shape's pixel coordinates, applies a custom DPI factor, converts the values to millimeters, and stores them in designated cells. | Explain how to adjust the pixel‑to‑centimeter conversion factor for different screen resolutions when working with shape positions in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to create a workbook with Aspose.Cells for .NET, add a rectangle shape, read its absolute Left and Top pixel offsets, convert those offsets to centimeters (using a 96 DPI reference), write the formatted results into cells A1 and B1, and save the workbook.
class ShapePositionToCell
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a sample rectangle shape (you can replace this with your own shape)
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 100, 100, 100);

        // Retrieve the shape's absolute position in pixels
        int leftPixels = shape.Left;   // horizontal offset from the left column
        int topPixels = shape.Top;     // vertical offset from the top row

        // Convert pixels to centimeters (assuming 96 DPI: 1 inch = 96 pixels, 1 inch = 2.54 cm)
        const double cmPerPixel = 2.54 / 96.0;
        double leftCm = leftPixels * cmPerPixel;
        double topCm = topPixels * cmPerPixel;

        // Store the converted values in cells
        worksheet.Cells["A1"].PutValue($"Left (cm): {leftCm:F2}");
        worksheet.Cells["B1"].PutValue($"Top (cm): {topCm:F2}");

        // Save the workbook
        workbook.Save("ShapePosition.xlsx");
    }
}
