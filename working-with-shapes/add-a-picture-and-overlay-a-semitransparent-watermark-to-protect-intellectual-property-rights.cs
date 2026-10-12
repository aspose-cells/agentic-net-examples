// Title: Add a JPEG image to an Excel worksheet and overlay a centered 45‑degree semi‑transparent "CONFIDENTIAL" text watermark using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that inserts a JPEG picture at cell B2, resizes it to 300 × 200 pixels, and overlays a 45‑degree 50% transparent "CONFIDENTIAL" text shape behind the picture. | Write a method that calculates the total worksheet width and height in pixels and positions a transparent text watermark at the sheet center using Aspose.Cells. | Provide a C# snippet that sets the Z‑order of a picture and a text shape so the watermark appears behind the image, then saves the workbook to a given file path.
// Common Searches: how to insert an image and a text watermark in the same Excel sheet using Aspose.Cells C# | Aspose.Cells C# place semi transparent WordArt watermark behind a picture | center a watermark on an Excel worksheet programmatically with Aspose.Cells | set ZOrderPosition for shapes in Aspose.Cells to control layering | calculate worksheet pixel dimensions Aspose.Cells C# for positioning shapes
// Tags: add picture to worksheet Aspose.Cells | semi transparent WordArt watermark Aspose.Cells | center watermark on Excel sheet C# | shape ZOrder positioning Aspose.Cells | save workbook with image and watermark .NET

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Creates a new workbook, inserts a JPEG at cell B2, adds a 45‑degree 50% transparent "CONFIDENTIAL" text watermark centered on the sheet behind the image, adjusts Z‑order, and saves the file.
class AddPictureWithWatermark
{
    static void Main()
    {
        try
        {
            // Path to the picture to be added
            string picturePath = @"C:\Images\sample.jpg";

            // Verify that the picture file exists
            if (!File.Exists(picturePath))
                throw new FileNotFoundException("Picture file not found.", picturePath);

            // Initialize a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add the picture to the worksheet at cell B2 (row 1, column 1)
            int pictureIndex = sheet.Pictures.Add(1, 1, picturePath);
            Picture picture = sheet.Pictures[pictureIndex];
            picture.Width = 300;
            picture.Height = 200;

            // Add a semi‑transparent text watermark using WordArt
            Shape watermark = sheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1,   // preset effect
                "CONFIDENTIAL",                    // watermark text
                "Arial",                           // font name
                48,                                // font size
                false,                             // bold
                false,                             // italic
                0, 0,                              // upper‑left cell (row, column)
                0, 0,                              // lower‑right cell (row, column)
                0, 0);                             // offsets

            // Set placement, rotation and transparency
            watermark.Placement = PlacementType.FreeFloating;
            watermark.RotationAngle = -45;
            watermark.Fill.Transparency = 0.5; // 50% transparent
            watermark.Font.Color = Color.Gray;

            // Position the watermark roughly at the center of the sheet
            int totalWidth = 0;
            for (int col = 0; col <= sheet.Cells.MaxColumn; col++)
                totalWidth += sheet.Cells.GetColumnWidthPixel(col);

            int totalHeight = 0;
            for (int row = 0; row <= sheet.Cells.MaxRow; row++)
                totalHeight += sheet.Cells.GetRowHeightPixel(row);

            watermark.Left = totalWidth / 2 - watermark.Width / 2;
            watermark.Top = totalHeight / 2 - watermark.Height / 2;

            // Ensure the watermark is behind the picture by adjusting Z‑order
            // Lower ZOrderPosition means farther back
            watermark.ZOrderPosition = 0;
            picture.ZOrderPosition = 1;

            // Prepare output path and ensure directory exists
            string outputPath = @"C:\Output\WorkbookWithPictureAndWatermark.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook to a file
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
