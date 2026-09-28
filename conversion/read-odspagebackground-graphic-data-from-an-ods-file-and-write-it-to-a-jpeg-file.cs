// Title: Extract the graphic page background from an ODS file and save it as a JPEG using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an ODS workbook with Aspose.Cells, reads the ODSPageBackground graphic data from the first worksheet, and writes the bytes to a JPEG file. | Modify the extraction routine to output the ODS page background as a PNG file, including error handling for missing or non‑graphic backgrounds. | Create a C# program that loops through all worksheets in an ODS workbook, extracts each ODSPageBackground graphic, and saves them as separate JPEG files named after the worksheet.
// Common Searches: C# Aspose.Cells extract ODS page background image to JPEG | How to save ODS worksheet background graphic as an image file in .NET | Aspose.Cells ODSPageBackground graphic data conversion example | Export ODS page background to PNG using Aspose.Cells C# | Batch extract page backgrounds from all sheets in an ODS file with Aspose.Cells
// Tags: ODSPageBackground graphic extraction to JPEG | Aspose.Cells save ODS background as image | C# write raw ODSPageBackground bytes to file | convert ODS page background to PNG | iterate worksheets export ODS backgrounds

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Ods;

// The sample loads an ODS workbook, accesses the first worksheet's ODSPageBackground, checks that it contains graphic data, and writes the raw bytes directly to a JPEG file.
class OdsPageBackgroundToJpeg
{
    static void Main()
    {
        try
        {
            // Path to the source ODS file
            string odsPath = "input.ods";

            // Path for the output JPEG file
            string jpegPath = "background.jpg";

            // Verify that the input file exists
            if (!File.Exists(odsPath))
            {
                Console.WriteLine($"Input file not found: {odsPath}");
                return;
            }

            // Load the ODS workbook
            Workbook workbook = new Workbook(odsPath);

            // Access the first worksheet's page setup
            PageSetup pageSetup = workbook.Worksheets[0].PageSetup;

            // Get the ODS page background object
            OdsPageBackground background = pageSetup.ODSPageBackground;

            // Ensure the background type is graphic and contains data
            if (background.Type == OdsPageBackgroundType.Graphic &&
                background.GraphicData != null &&
                background.GraphicData.Length > 0)
            {
                // Save the raw graphic data directly as a JPEG file
                File.WriteAllBytes(jpegPath, background.GraphicData);
                Console.WriteLine($"Graphic background extracted and saved to {jpegPath}");
            }
            else
            {
                Console.WriteLine("No graphic background found in the ODS file.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
