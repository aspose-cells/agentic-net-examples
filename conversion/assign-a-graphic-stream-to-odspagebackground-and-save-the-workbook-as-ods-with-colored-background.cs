// Title: Set a tiled PNG graphic with light‑green fill as the ODS page background using Aspose.Cells for .NET
// AI Prompts: Write C# code that reads a PNG file into a byte array, assigns it to ODSPageBackground.GraphicData, sets the background type to Graphic, chooses tiled center positioning, applies a light‑green background color, and saves the workbook as an .ods file with Aspose.Cells. | Create a reusable C# method that takes an image Stream, configures the ODS page background to use the image as a tiled graphic with a specified fill color, and writes the workbook to an ODS file.
// Common Searches: how to set a graphic background image for ODS files using Aspose.Cells .NET | Aspose.Cells ODSPageBackground tile image and color example | C# add background picture and fill color to ODS workbook | set ODS page background graphic from file stream Aspose.Cells
// Tags: Aspose.Cells ODS graphic background tiling | C# set ODS document page image | add background color to ODS file with Aspose.Cells | load PNG data into workbook page background | export workbook to ODS with custom background

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Ods;

namespace AsposeCellsExamples
{
    // The example creates a new Workbook, accesses the first worksheet's PageSetup, configures ODSPageBackground to use a tiled PNG graphic centered on the page, applies a light‑green fill color, and saves the result as an ODS file.
    public class OdsPageBackgroundGraphicDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Get the PageSetup object
                PageSetup pageSetup = worksheet.PageSetup;

                // Access the ODS page background
                OdsPageBackground background = pageSetup.ODSPageBackground;

                // Set the background type to Graphic
                background.Type = OdsPageBackgroundType.Graphic;

                // Load image data into a byte array (replace with actual image path)
                string imagePath = "background.png";
                if (File.Exists(imagePath))
                {
                    background.GraphicData = File.ReadAllBytes(imagePath);
                }
                else
                {
                    Console.WriteLine("Image file not found: " + imagePath);
                    return;
                }

                // Set graphic formatting options
                background.GraphicType = OdsPageBackgroundGraphicType.Tile;
                background.GraphicPositionType = OdsPageBackgroundGraphicPositionType.CenterCenter;

                // Also set a background color (will be visible if graphic has transparency)
                background.Color = Color.LightGreen;

                // Save the workbook as ODS
                string outputPath = "OdsPageBackgroundWithGraphic.ods";
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved successfully to: " + outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
