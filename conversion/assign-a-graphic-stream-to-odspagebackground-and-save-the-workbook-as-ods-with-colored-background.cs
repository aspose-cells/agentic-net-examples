// Title: C# – Assign a graphic stream to ODS page background and save workbook with solid color using AspNet Cells
// Description: Demonstrates how to create a Workbook, set a LightGreen page color, load an image via FileStream into a byte array, apply it as a tiled graphic background on the ODS page, and export the file as an ODS document with Aspose.Cells for .NET.
// Keywords: Aspose.Cells ODS background image | C# OdsPageBackground graphic stream | set ODS page color Aspose.Cells | tile background image ODS | save workbook as ODS .NET
// Common Searches: how to add image background to ODS with Aspose.Cells C# | assign graphic data to OdsPageBackground in .NET | ODS page background color and tiled image example | export Excel to ODS with custom background
// Developer Intent: Apply a bitmap graphic (from a stream) as the tiled page background of an ODS worksheet while also defining a solid fill color, then generate the ODS file.
// Use Cases: Design printable ODS reports that include a company logo repeated across the page. | Create template spreadsheets with a corporate color scheme and watermark background. | Generate ODS files for distribution where a light‑green fill and centered graphic improve visual branding.
// AI Prompts: Provide C# code that reads an image file into a byte array and sets OdsPageBackground.GraphicData in Aspose.Cells. | Show an example of configuring OdsPageBackground.Type, GraphicType, and GraphicPositionType for a tiled background. | Explain how to combine a solid page color with a graphic background and save the result as an ODS file.

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Ods;

// Demonstrates how to create a Workbook, set a LightGreen page color, load an image via FileStream into a byte array, apply it as a tiled graphic background on the ODS page, and export the file as an ODS document with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Access the ODS page background object
        OdsPageBackground background = sheet.PageSetup.ODSPageBackground;

        // Set a solid background color (e.g., LightGreen)
        background.Color = Color.LightGreen;

        // Path to the image that will be used as a graphic background
        string imagePath = "background.png"; // Replace with your image file path

        if (File.Exists(imagePath))
        {
            // Load the image into a byte array using a stream
            byte[] imageData;
            using (FileStream fileStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
            using (MemoryStream memoryStream = new MemoryStream())
            {
                fileStream.CopyTo(memoryStream);
                imageData = memoryStream.ToArray();
            }

            // Configure the page background to use the graphic data
            background.Type = OdsPageBackgroundType.Graphic;          // Use graphic background
            background.GraphicData = imageData;                      // Assign the image bytes
            background.GraphicType = OdsPageBackgroundGraphicType.Tile; // Tile the image
            background.GraphicPositionType = OdsPageBackgroundGraphicPositionType.CenterCenter; // Center the tiles
        }

        // Save the workbook as an ODS file
        workbook.Save("WorkbookWithGraphicBackground.ods");
    }
}
