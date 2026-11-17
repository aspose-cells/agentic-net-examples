// Title: Add image background and semi‑transparent text watermarks to every worksheet in an Excel file with Aspose.Cells for .NET
// Common Searches: how to apply a background image watermark to all worksheets using Aspose.Cells C# | Aspose.Cells C# add semi transparent text watermark over worksheet background image | set worksheet background picture and foreground text shape as watermark in .NET | C# loop through worksheets to add image and text watermarks with Aspose.Cells | Aspose.Cells example for combined image and text watermarks in Excel workbook
// Tags: worksheet background image watermark Aspose.Cells | foreground text shape watermark C# | apply combined image and text watermarks Excel Aspose.Cells | iterate worksheets add watermarks Aspose.Cells | semi transparent text watermark Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, reads a PNG once, then iterates through each worksheet to set the BackgroundImage property for an image watermark and adds a free‑floating text shape with semi‑transparent fill as a foreground watermark, finally saving the watermarked workbook.
class ExcelWatermarkExample
{
    static void Main()
    {
        // Paths for the source workbook, watermark image, and output workbook
        string workbookPath = @"C:\Docs\source.xlsx";
        string imagePath = @"C:\Docs\watermark.png";
        string outputPath = @"C:\Docs\watermarked.xlsx";

        try
        {
            // Verify that required files exist
            if (!File.Exists(workbookPath))
                throw new FileNotFoundException("Source workbook not found.", workbookPath);
            if (!File.Exists(imagePath))
                throw new FileNotFoundException("Watermark image not found.", imagePath);

            // Load the Excel workbook
            Workbook workbook = new Workbook(workbookPath);

            // Read image bytes once (used for all worksheets)
            byte[] imageBytes = File.ReadAllBytes(imagePath);

            // Apply watermarks to each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // ----- Image watermark (background) -----
                // Set the worksheet background image (appears behind cell content)
                sheet.BackgroundImage = imageBytes;

                // ----- Text watermark (foreground) -----
                // Add a text shape that will act as a watermark
                Shape textShape = sheet.Shapes.AddTextEffect(
                    MsoPresetTextEffect.TextEffect1,
                    "CONFIDENTIAL",
                    "Arial",
                    72,
                    true,
                    false,
                    0,   // left
                    0,   // top
                    0,   // width (auto)
                    0,   // height (auto)
                    45,  // rotation angle
                    0);  // text effect shape (default)

                // Position the shape (free floating)
                textShape.Placement = PlacementType.FreeFloating;
                textShape.Top = 0;
                textShape.Left = 0;

                // Set semi‑transparent fill (0 = opaque, 1 = fully transparent)
                textShape.Fill.Transparency = 0.5;

                // Hide the outline (set line weight to 0)
                textShape.Line.Weight = 0;
            }

            // Ensure output directory exists
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine("Watermarked workbook saved to: " + outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
