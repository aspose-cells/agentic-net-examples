// Title: Extract ODS Page Background Image to JPEG with Aspose.Cells for .NET (C#)
// Description: Loads an ODS workbook using Aspose.Cells, accesses a worksheet's OdsPageBackground, verifies the graphic type, and writes the raw image bytes to a JPEG file. Includes error handling for missing files or non‑graphic backgrounds.
// Keywords: Aspose.Cells ODS background extraction | C# OdsPageBackground graphic data | save ODS page background as JPEG | extract ODS background image .NET | Aspose.Cells read ODS page background | ODS to JPEG conversion C# | Aspose.Cells ODSPageBackgroundType Graphic
// Common Searches: how to extract background image from ODS using Aspose.Cells | C# code to save ODS page background as JPEG | Aspose.Cells OdsPageBackground example | read ODS page background graphic data in .NET | convert ODS background to image with Aspose
// Developer Intent: Retrieve the graphic page background from an ODS worksheet and write it to a JPEG file using Aspose.Cells.
// Use Cases: Create thumbnail previews of ODS sheets by extracting embedded background graphics. | Migrate ODS background images to other formats or content management systems. | Validate the presence of a background image before processing ODS documents.
// AI Prompts: Generate C# code that extracts an ODS worksheet's page background and saves it as PNG using Aspose.Cells. | Explain how to handle OdsPageBackground types other than Graphic when extracting images. | Show how to extract the background from a specific worksheet index and return it as a MemoryStream.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Ods;

// Loads an ODS workbook using Aspose.Cells, accesses a worksheet's OdsPageBackground, verifies the graphic type, and writes the raw image bytes to a JPEG file. Includes error handling for missing files or non‑graphic backgrounds.
class ExtractOdsPageBackground
{
    static void Main()
    {
        // Path to the source ODS file
        string odsFilePath = "input.ods";

        // Path where the extracted JPEG will be saved
        string jpegOutputPath = "background.jpg";

        try
        {
            // Verify that the input ODS file exists
            if (!File.Exists(odsFilePath))
            {
                Console.WriteLine($"Input file not found: {odsFilePath}");
                return;
            }

            // Load the ODS workbook
            Workbook workbook = new Workbook(odsFilePath);

            // Access the first worksheet (you can choose any worksheet)
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the ODS page background object
            OdsPageBackground pageBackground = worksheet.PageSetup.ODSPageBackground;

            // Verify that the background is a graphic and that data exists
            if (pageBackground.Type == OdsPageBackgroundType.Graphic &&
                pageBackground.GraphicData != null &&
                pageBackground.GraphicData.Length > 0)
            {
                // Write the raw graphic data directly to a file
                File.WriteAllBytes(jpegOutputPath, pageBackground.GraphicData);
                Console.WriteLine($"Background image extracted successfully to: {jpegOutputPath}");
            }
            else
            {
                Console.WriteLine("The ODS file does not contain a graphic page background.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
