// Title: Extract ODS page graphic background to PNG with Aspose.Cells for .NET
// Description: C# code that loads an ODS workbook via Aspose.Cells, reads the first worksheet's OdsPageBackground, confirms it is a graphic, obtains the raw image bytes, and writes them to a PNG file.
// Keywords: Aspose.Cells | C# | .NET | ODS | extract page background | OdsPageBackground | graphic background | PNG export | convert ODS to image | Aspose.Cells ODS background
// Common Searches: Aspose.Cells extract ODS background image | Save ODS page graphic as PNG C# | How to get OdsPageBackground bytes | Convert ODS worksheet background to PNG | C# code for ODS background extraction
// Developer Intent: Retrieve the graphic page background from an ODS file and write it to a PNG file.
// Use Cases: Reuse a logo embedded as an ODS page background by extracting it as a standalone PNG. | Create web‑ready thumbnails of ODS worksheets by exporting their background graphics. | Validate and archive background images from ODS templates before further processing.
// AI Prompts: Generate C# code that loads an ODS workbook with Aspose.Cells, checks for a graphic page background, and saves it as a PNG file. | Explain the role of OdsPageBackground.Type and OdsPageBackground.GraphicData when extracting background images from ODS files. | Provide robust error‑handling patterns for reading ODS page backgrounds and writing the image data to disk in .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Ods;

namespace AsposeCellsExamples
{
    // C# code that loads an ODS workbook via Aspose.Cells, reads the first worksheet's OdsPageBackground, confirms it is a graphic, obtains the raw image bytes, and writes them to a PNG file.
    public class ExtractOdsPageBackground
    {
        public static void Run()
        {
            // Path to the source ODS file
            string odsPath = "input.ods";

            // Verify that the ODS file exists
            if (!File.Exists(odsPath))
            {
                Console.WriteLine($"Input file not found: {odsPath}");
                return;
            }

            try
            {
                // Load the ODS workbook
                Workbook workbook = new Workbook(odsPath);

                // Access the first worksheet (index 0)
                Worksheet sheet = workbook.Worksheets[0];

                // Get the ODS page background settings
                OdsPageBackground background = sheet.PageSetup.ODSPageBackground;

                // Verify that the background is of graphic type
                if (background.Type == OdsPageBackgroundType.Graphic && background.GraphicData != null)
                {
                    // Retrieve the raw graphic data (bytes)
                    byte[] imageBytes = background.GraphicData;

                    // Define the output PNG file path
                    string pngPath = "background.png";

                    // Write the bytes to a PNG file
                    File.WriteAllBytes(pngPath, imageBytes);

                    Console.WriteLine($"Graphic background extracted and saved to: {pngPath}");
                }
                else
                {
                    Console.WriteLine("The ODS file does not contain a graphic page background.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the ODS file: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                ExtractOdsPageBackground.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
