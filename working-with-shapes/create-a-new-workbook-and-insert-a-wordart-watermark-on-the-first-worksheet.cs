using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class WordArtWatermarkDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the shape collection of the worksheet
            ShapeCollection shapes = sheet.Shapes;

            // Add a WordArt shape that will serve as a watermark
            // Parameters: style, text, topRow, top (pixel offset), leftColumn, left (pixel offset), height, width
            Shape wordArt = shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1,   // preset style
                "CONFIDENTIAL",                     // watermark text
                0, 0,                               // top row and pixel offset
                0, 0,                               // left column and pixel offset
                100, 400);                          // height and width in pixels

            // Rotate the WordArt to give a typical watermark appearance
            wordArt.RotationAngle = 45; // degrees

            // Make the WordArt semi‑transparent
            wordArt.Fill.Transparency = 0.7; // 70% transparent

            // Hide the outline of the shape by making it fully transparent
            wordArt.Line.Transparency = 1.0; // 100% transparent (invisible)

            // Define output file path
            string outputPath = "WordArtWatermark.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}