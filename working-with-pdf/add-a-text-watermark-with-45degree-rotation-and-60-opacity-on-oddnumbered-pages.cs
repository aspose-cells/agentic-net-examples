// Title: Add a 45-degree rotated text watermark with 60% opacity to an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# how to add a diagonal text watermark with custom opacity | set rotation angle and transparency for a shape watermark in an Excel file using Aspose.Cells | C# example to create a confidential watermark on a worksheet with Aspose.Cells | apply 60% opacity to a text effect shape in Aspose.Cells workbook | rotate and send watermark shape to back of cells using Aspose.Cells .NET
// Tags: Aspose.Cells rotate watermark shape | Aspose.Cells set shape transparency | C# add text effect watermark Excel | Excel worksheet diagonal watermark Aspose.Cells | Aspose.Cells watermark Z-order positioning

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The sample creates a new Workbook, adds a text effect shape with the text "Confidential" to the first worksheet, rotates it 45°, sets its fill transparency to 0.4 (60% opacity), moves the shape behind cell content, and saves the file as WatermarkedWorkbook.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add a text shape that will act as a watermark
                // Parameters: preset effect, text, font name, size, bold, italic,
                // left, top, width, height, leftColumn, topRow
                Shape watermarkShape = sheet.Shapes.AddTextEffect(
                    MsoPresetTextEffect.TextEffect1,
                    "Confidential",
                    "Arial",
                    72,
                    false,
                    false,
                    0,      // left
                    0,      // top
                    500,    // width
                    200,    // height
                    0,      // leftColumn
                    0);     // topRow

                // Rotate the shape to achieve a diagonal watermark
                watermarkShape.RotationAngle = 45;

                // Set transparency (0 = opaque, 1 = fully transparent)
                watermarkShape.Fill.Transparency = 0.4;

                // Send the shape to the back so it appears behind cell content
                watermarkShape.ZOrderPosition = 0;

                // Save the workbook
                string outputPath = "WatermarkedWorkbook.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
