using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Ods;
using Aspose.Cells.Rendering;

class VerifyOdsBackgroundInPdf
{
    static void Main()
    {
        // Path to the image that will be used as the ODS page background
        string imagePath = "background.png"; // Ensure this file exists in the executable directory

        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Access the ODS page background through the worksheet's PageSetup
        OdsPageBackground odsBackground = sheet.PageSetup.ODSPageBackground;

        // Configure the background to use a graphic
        odsBackground.Type = OdsPageBackgroundType.Graphic;
        odsBackground.GraphicType = OdsPageBackgroundGraphicType.Tile;               // Repeat the image
        odsBackground.GraphicPositionType = OdsPageBackgroundGraphicPositionType.CenterCenter;

        // Load the image bytes and assign them to the background
        if (File.Exists(imagePath))
        {
            odsBackground.GraphicData = File.ReadAllBytes(imagePath);
        }
        else
        {
            Console.WriteLine($"Image file not found: {imagePath}");
            return;
        }

        // Optional: add some visible content to the worksheet to see the background effect
        sheet.Cells["A1"].PutValue("Sample data with ODS background graphic");

        // Save the workbook as ODS (optional, for manual verification)
        workbook.Save("SampleWithBackground.ods");

        // Prepare PDF save options (no special settings required for background)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Export the workbook to PDF; the ODS background graphic should be rendered in the PDF
        workbook.Save("SampleWithBackground.pdf", pdfOptions);

        Console.WriteLine("PDF generated. Verify that the background graphic appears in the PDF file.");
    }
}