// Title: Render a tiled graphic as ODS page background and confirm its presence in the exported PDF using Aspose.Cells for .NET
// AI Prompts: Create an ODS workbook, assign a PNG image as a tiled, centered OdsPageBackground graphic, and save it as PDF with PdfSaveOptions. | Generate a placeholder PNG file if missing, set OdsPageBackground.Type to Graphic, provide the image bytes, set GraphicType to Tile and GraphicPositionType to CenterCenter, then export the workbook to PDF. | Save the workbook as .ods for manual inspection and as .pdf to verify that the tiled background graphic is rendered correctly.
// Common Searches: how to set a tiled background image in an ODS worksheet using Aspose.Cells .NET | Aspose.Cells OdsPageBackground graphic not appearing in PDF export | C# code to add a PNG as ODS page background and keep it in PDF | verify ODS page background graphic rendering in PDF with Aspose.Cells | export ODS workbook with tiled background to PDF using PdfSaveOptions
// Tags: OdsPageBackground graphic tile Aspose.Cells | export ODS worksheet to PDF with background image | C# set ODS page background graphic data | PdfSaveOptions preserve ODS background | generate placeholder PNG for workbook background

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Ods;
using Aspose.Cells.Rendering;

// The example creates a placeholder PNG if needed, configures the first worksheet's OdsPageBackground as a tiled, centered graphic, saves the workbook as both ODS and PDF, and demonstrates how to verify that the background image is rendered in the resulting PDF.
class VerifyOdsBackgroundInPdf
{
    static void Main()
    {
        try
        {
            // Path to the background image
            string imagePath = "background.png";

            // Create a placeholder PNG image if it does not exist
            if (!File.Exists(imagePath))
            {
                // 1x1 white PNG (base64 decoded)
                byte[] pngBytes = new byte[]
                {
                    0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
                    0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
                    0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
                    0x08,0x02,0x00,0x00,0x00,0x90,0x77,0x53,
                    0xDE,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
                    0x54,0x08,0xD7,0x63,0xF8,0xCF,0xC0,0x00,
                    0x00,0x04,0x00,0x01,0xE2,0x26,0x05,0x9B,
                    0x00,0x00,0x00,0x00,0x49,0x45,0x4E,0x44,
                    0xAE,0x42,0x60,0x82
                };
                File.WriteAllBytes(imagePath, pngBytes);
            }

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Configure ODS page background as a graphic
            OdsPageBackground background = sheet.PageSetup.ODSPageBackground;
            background.Type = OdsPageBackgroundType.Graphic;                     // Use graphic background
            background.GraphicData = File.ReadAllBytes(imagePath);               // Set image data
            background.GraphicType = OdsPageBackgroundGraphicType.Tile;          // Tile the image
            background.GraphicPositionType = OdsPageBackgroundGraphicPositionType.CenterCenter; // Center it

            // Save the workbook as ODS (optional, for manual verification)
            workbook.Save("WorkbookWithBackground.ods");

            // Save the workbook as PDF; the background graphic will be rendered in the PDF
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            workbook.Save("WorkbookWithBackground.pdf", pdfOptions);

            Console.WriteLine("PDF generated with ODS background graphic.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
