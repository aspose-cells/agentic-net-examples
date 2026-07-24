// Title: Extract embedded ODS page background as PNG with Aspose.Cells for .NET
// Description: Loads an ODS workbook, accesses the first worksheet's PageSetup, verifies an embedded OdsPageBackground of type Graphic, reads the raw GraphicData bytes, and writes them to a PNG file.
// Keywords: Aspose.Cells | .NET | ODS background extraction | OdsPageBackground | GraphicData | save as PNG | embedded graphic | page setup | C# example | convert ODS to image
// Common Searches: Aspose.Cells extract ODS background C# | save ODS page background as PNG | OdsPageBackground GraphicData example | how to get embedded graphic from ODS file | convert ODS background to image using .NET
// Developer Intent: Retrieve the embedded graphic background from an ODS file and write it to a PNG image.
// Use Cases: Reuse the original ODS background image in other documents or web assets. | Generate thumbnail previews of ODS sheets for a document management system. | Validate presence of an embedded background before further ODS processing.
// AI Prompts: Write C# code that extracts an ODS page background with Aspose.Cells and saves it as JPEG. | Explain how to handle linked ODS backgrounds (IsLink = true) and download the external image. | Provide a script to batch‑process multiple ODS files, extracting each background to a separate folder.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Ods;

namespace AsposeCellsExamples
{
    // Loads an ODS workbook, accesses the first worksheet's PageSetup, verifies an embedded OdsPageBackground of type Graphic, reads the raw GraphicData bytes, and writes them to a PNG file.
    class ExtractOdsBackground
    {
        static void Main()
        {
            // Path to the source ODS file
            string odsPath = "input.ods";

            // Load the ODS workbook
            Workbook workbook = new Workbook(odsPath);

            // Access the first worksheet's page setup
            Worksheet sheet = workbook.Worksheets[0];
            PageSetup pageSetup = sheet.PageSetup;

            // Get the ODS page background object
            OdsPageBackground background = pageSetup.ODSPageBackground;

            // Verify that a graphic background is present and it is embedded (not linked)
            if (background.Type == OdsPageBackgroundType.Graphic && !background.IsLink)
            {
                // Retrieve the raw graphic data (byte array)
                byte[] imageData = background.GraphicData;

                // Save the graphic data as a PNG file
                string pngPath = "background.png";
                File.WriteAllBytes(pngPath, imageData);

                Console.WriteLine($"Graphic background extracted and saved to: {pngPath}");
            }
            else
            {
                Console.WriteLine("No embedded graphic background found in the ODS file.");
            }
        }
    }
}
