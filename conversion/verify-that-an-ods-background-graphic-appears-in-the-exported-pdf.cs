// Title: C# – Verify ODS Page Background Graphic Appears in PDF Using Aspose.Cells
// Description: Creates a workbook, sets the first worksheet’s ODS page background to a PNG graphic, configures the image to tile and center, and saves the workbook as PDF with PdfSaveOptions, showing how to ensure the background graphic is rendered in the exported PDF.
// Keywords: Aspose.Cells | C# | .NET | ODS page background | background graphic | PDF export | PdfSaveOptions | OdsPageBackground | tile background | centered background | verify PDF content
// Common Searches: add graphic background to ODS worksheet and export to PDF Aspose.Cells | check ODS page background in generated PDF C# | Aspose.Cells ODSPageBackground graphic example | C# export ODS workbook with background image to PDF
// Developer Intent: Confirm that the ODS worksheet background image is included in the PDF produced by Aspose.Cells.
// Use Cases: Apply a PNG image as a tiled, centered background to a worksheet before PDF conversion. | Automate validation that the exported PDF contains the expected background graphic. | Reuse the same background graphic across multiple worksheets in a single PDF document.
// AI Prompts: Generate C# code that loads a PNG, sets it as an ODS page background with tiling and centering, and saves the workbook to PDF using Aspose.Cells. | Explain how to programmatically verify that the background graphic is present in the PDF created by Aspose.Cells. | Provide step‑by‑step instructions for configuring OdsPageBackgroundGraphicType.Tile and OdsPageBackgroundGraphicPositionType.CenterCenter before exporting to PDF.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Ods;
using Aspose.Cells.Rendering;

namespace OdsBackgroundPdfVerification
{
    // Creates a workbook, sets the first worksheet’s ODS page background to a PNG graphic, configures the image to tile and center, and saves the workbook as PDF with PdfSaveOptions, showing how to ensure the background graphic is rendered in the exported PDF.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Access the ODS page background of the worksheet
            OdsPageBackground background = sheet.PageSetup.ODSPageBackground;

            // Set the background type to Graphic
            background.Type = OdsPageBackgroundType.Graphic;

            // Load image data to be used as background graphic
            string imagePath = "background.png"; // Ensure this file exists in the executable directory
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }
            background.GraphicData = File.ReadAllBytes(imagePath);

            // Configure graphic display options
            background.GraphicType = OdsPageBackgroundGraphicType.Tile;
            background.GraphicPositionType = OdsPageBackgroundGraphicPositionType.CenterCenter;

            // Save the workbook as PDF; the ODS background graphic will be rendered in the PDF
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Ensure the background graphic is included; no special option needed beyond default
                // Additional options can be set here if required
            };

            string pdfPath = "WorkbookWithBackground.pdf";
            workbook.Save(pdfPath, pdfOptions);

            Console.WriteLine($"PDF saved successfully to {pdfPath}");
        }
    }
}
