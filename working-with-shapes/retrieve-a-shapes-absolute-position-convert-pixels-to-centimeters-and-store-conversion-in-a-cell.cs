// Title: C# – Retrieve Shape Position, Convert Pixels to Centimeters, and Store in Cells with Aspose.Cells
// Description: Creates a workbook, adds a rectangle shape, reads its absolute Left and Top pixel coordinates, converts them to centimeters using the 2.54 cm/96 DPI factor, writes the metric values to cells, and saves the file.
// Keywords: Aspose.Cells shape position | pixel to centimeter conversion | C# Aspose.Cells shape coordinates | absolute shape location Excel | store shape metrics in worksheet
// Common Searches: Aspose.Cells get shape absolute position C# | convert shape pixels to centimeters Aspose.Cells | write shape coordinates to Excel cells | pixel to cm factor Aspose.Cells shapes | Aspose.Cells shape layout measurement
// Developer Intent: Obtain a shape's pixel‑based Left and Top values, translate them into centimeters, and record the results in specific worksheet cells.
// Use Cases: Generate a printable layout report that lists each shape's physical position in centimeters. | Validate diagram alignment by comparing metric coordinates against design specifications. | Migrate legacy pixel‑based drawings to metric standards for downstream processing.
// AI Prompts: Show C# code using Aspose.Cells to read a shape's Left and Top pixel values, convert them to centimeters, and write the results to cells A1:B2. | Explain how to calculate the pixel‑to‑centimeter conversion factor for Aspose.Cells shapes and apply it to multiple shapes. | Provide a loop that iterates over all shapes on a worksheet, converts their positions to centimeters, and creates a summary table in the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapePositionDemo
{
    // Creates a workbook, adds a rectangle shape, reads its absolute Left and Top pixel coordinates, converts them to centimeters using the 2.54 cm/96 DPI factor, writes the metric values to cells, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, upper left pixel offset X, upper left pixel offset Y, width in pixels, height in pixels
            Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 150, 80);

            // Retrieve the shape's absolute position in pixels (Left and Top properties are in pixels)
            int leftPixels = shape.Left;
            int topPixels = shape.Top;

            // Conversion factor: 1 pixel = 2.54 cm / 96 DPI
            const double pixelToCm = 2.54 / 96.0;

            // Convert pixel values to centimeters
            double leftCm = leftPixels * pixelToCm;
            double topCm = topPixels * pixelToCm;

            // Store the converted values in cells
            worksheet.Cells["A1"].PutValue("Left (cm)");
            worksheet.Cells["B1"].PutValue(leftCm);
            worksheet.Cells["A2"].PutValue("Top (cm)");
            worksheet.Cells["B2"].PutValue(topCm);

            // Save the workbook
            workbook.Save("ShapePositionInCm.xlsx");
        }
    }
}
