// Title: Extract ODS Page Background Image to JPEG with Aspose.Cells for .NET
// Description: Loads an ODS workbook, reads the first worksheet's ODSPageBackground graphic data, and writes the bytes to a JPEG file while handling missing files and absent backgrounds.
// Keywords: Aspose.Cells ODS page background | extract ODS background image | save ODS graphic as JPEG | OdsPageBackground GraphicData .NET | convert ODS background to image
// Common Searches: Aspose.Cells read ODS page background image | C# extract ODS worksheet background to JPEG | How to get OdsPageBackground graphic data | Save ODS background as picture using .NET | Extract ODS background image with Aspose
// Developer Intent: Retrieve the background graphic from an ODS sheet and store it as a JPEG file.
// Use Cases: Create thumbnails of ODS worksheets by extracting their background images. | Archive original ODS page graphics for documentation or compliance. | Generate reports that embed the exact ODS background as a standalone image.
// AI Prompts: Write C# code that uses Aspose.Cells to read OdsPageBackground.GraphicData from an ODS file and save it as a JPEG, including error handling for missing backgrounds. | Explain how to check for null or empty GraphicData before writing to disk and how to extend the code to support PNG or BMP output. | Show a loop that processes every worksheet in a workbook, extracting each ODS page background to separate image files with unique names.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Ods;

// Loads an ODS workbook, reads the first worksheet's ODSPageBackground graphic data, and writes the bytes to a JPEG file while handling missing files and absent backgrounds.
class ExtractOdsPageBackground
{
    static void Main()
    {
        try
        {
            // Input ODS file path
            string odsPath = "input.ods";

            // Output JPEG file path
            string jpegPath = "background.jpg";

            // Verify input file exists
            if (!File.Exists(odsPath))
            {
                Console.WriteLine($"Input file not found: {odsPath}");
                return;
            }

            // Load the ODS workbook
            Workbook workbook = new Workbook(odsPath);

            // Access the first worksheet's ODS page background
            OdsPageBackground background = workbook.Worksheets[0].PageSetup.ODSPageBackground;

            // Retrieve the graphic data (image bytes)
            byte[] graphicData = background?.GraphicData;

            if (graphicData == null || graphicData.Length == 0)
            {
                Console.WriteLine("No graphic background found in the ODS file.");
                return;
            }

            // Save the image bytes directly to a file
            File.WriteAllBytes(jpegPath, graphicData);

            Console.WriteLine($"Graphic background extracted to: {jpegPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
