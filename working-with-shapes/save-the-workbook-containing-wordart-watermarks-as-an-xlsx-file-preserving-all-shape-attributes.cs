// Title: Add a rotated semi‑transparent WordArt rectangle watermark to an Excel workbook and save as XLSX using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that creates a rectangle shape, assigns custom text, font, color, 80% fill transparency, -45° rotation, hides the border, and saves the workbook as an XLSX file. | Show how to programmatically verify and create the output directory before invoking Workbook.Save when generating an Excel file that contains a WordArt watermark using Aspose.Cells. | Demonstrate applying WordArt formatting (font name, size, style, color) and shape properties (transparency, rotation) to produce a diagonal watermark in an Excel worksheet with Aspose.Cells for .NET.
// Common Searches: aspnet add word art watermark to excel worksheet using aspose.cells | c# create rotated transparent shape as excel watermark with aspose.cells | save excel file with shape watermark ensuring output directory exists in .net | how to set shape fill transparency and rotation in aspose.cells c#
// Tags: Aspose.Cells add diagonal watermark shape | Aspose.Cells set shape fill transparency | Aspose.Cells hide shape line weight | Aspose.Cells save workbook as xlsx with shapes | C# create output folder before Workbook.Save

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // This example creates a new workbook, inserts a rectangular WordArt shape with the text "CONFIDENTIAL" (Arial, 72pt, red), applies 80% transparency, rotates it -45° for a diagonal effect, removes the border, ensures the target folder exists, and saves the file as WordArtWatermark.xlsx using Aspose.Cells for .NET.
class WordArtWatermarkExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape that will hold the watermark text
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width (all in points)
            Shape watermark = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // Rectangle shape
                0,   // Upper left row
                0,   // Upper left column
                0,   // Top offset (points)
                0,   // Left offset (points)
                200, // Height (points)
                500  // Width (points)
            );

            // Set the watermark text and its formatting
            watermark.Text = "CONFIDENTIAL";
            watermark.Font.Name = "Arial";
            watermark.Font.Size = 72;
            watermark.Font.IsBold = false;
            watermark.Font.IsItalic = false;
            watermark.Font.Color = Color.Red;

            // Make the shape semi‑transparent
            watermark.Fill.Transparency = 0.8; // 80% transparent

            // Rotate the shape for diagonal appearance
            watermark.RotationAngle = -45;

            // Hide the shape border (if supported)
            // Note: Some versions of Aspose.Cells may not expose a direct property to hide the border.
            // The line can be made invisible by setting its weight to 0.
            watermark.Line.Weight = 0;

            // Define output file path
            string outputPath = "WordArtWatermark.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
