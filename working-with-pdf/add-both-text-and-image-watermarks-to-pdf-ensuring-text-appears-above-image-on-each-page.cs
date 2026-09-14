// Title: Add a PNG background image and diagonal CONFIDENTIAL text watermark (front) to a PDF generated from an Aspose.Cells workbook in C#
// AI Prompts: Generate C# code using Aspose.Cells to place a PNG image as a background watermark, then overlay a rotated "CONFIDENTIAL" text shape on top, and export the worksheet to PDF. | Modify the watermark example to set the image opacity to 30%, change the text font to Times New Roman, size 80, color red, and keep the text above the image before saving as PDF.
// Common Searches: how to add image watermark behind worksheet and text watermark in front using Aspose.Cells C# | Aspose.Cells set ZOrderPosition for picture and shape to control watermark layering | save workbook as PDF with diagonal text watermark over background image in C# | Aspose.Cells C# rotate text shape for watermark on PDF export | apply opacity to image watermark in Aspose.Cells before PDF conversion
// Tags: Aspose.Cells add image watermark behind worksheet | Aspose.Cells overlay text watermark on PDF | C# set ZOrderPosition shape ordering Aspose.Cells | Aspose.Cells rotate text shape watermark | Aspose.Cells export workbook to PDF with watermarks

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates creating a Workbook, inserting a PNG image as a background watermark, adding a rotated "CONFIDENTIAL" text shape in front, and saving the worksheet as a PDF using Aspose.Cells for .NET.
class WatermarkPdfExample
{
    static void Main()
    {
        Workbook workbook = null;

        // Create workbook and add watermarks
        try
        {
            workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data (optional)
            sheet.Cells["A1"].PutValue("Sample Data");
            sheet.Cells["A2"].PutValue(12345);

            // -------------------------------------------------
            // Add image watermark (placed behind the sheet)
            // -------------------------------------------------
            string imagePath = "watermarkImage.png";

            if (File.Exists(imagePath))
            {
                // Add the picture to the worksheet using the file path
                int pictureIndex = sheet.Pictures.Add(0, 0, imagePath);
                Picture picture = sheet.Pictures[pictureIndex];

                // Send the picture to the back so other objects appear above it
                picture.ZOrderPosition = 0; // 0 = back
            }
            else
            {
                Console.WriteLine($"Image file not found: {imagePath}. Skipping image watermark.");
            }

            // -------------------------------------------------
            // Add text watermark (placed above the image)
            // -------------------------------------------------
            Shape textShape = sheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1,
                "CONFIDENTIAL",
                "Arial",
                72,
                true,
                false,
                0,
                0,
                500,
                100,
                0,
                0);

            // Bring the text shape to the front
            textShape.ZOrderPosition = 1;

            // Optional: diagonal rotation
            textShape.RotationAngle = -45;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while creating the workbook or adding watermarks: {ex.Message}");
        }

        // Save the workbook as PDF
        try
        {
            if (workbook != null)
            {
                workbook.Save("WatermarkedDocument.pdf", SaveFormat.Pdf);
                Console.WriteLine("PDF saved successfully as WatermarkedDocument.pdf");
            }
            else
            {
                Console.WriteLine("Workbook was not created; PDF not saved.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save PDF: {ex.Message}");
        }
    }
}
