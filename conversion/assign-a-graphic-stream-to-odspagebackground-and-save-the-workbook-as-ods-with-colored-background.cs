// Title: Set a tiled graphic and background color for ODS worksheets using Aspose.Cells C#
// Description: Demonstrates how to create a workbook, access the ODSPageBackground of the first worksheet, assign a PNG image as a tiled graphic, set a background color, and save the file as an ODS document with Aspose.Cells for .NET.
// Keywords: Aspose.Cells ODS background image | C# OdsPageBackground graphic | tile graphic ODS worksheet | ODS page background color | save workbook as ODS | Aspose.Cells .NET example
// Common Searches: How to add an image background to an ODS file with Aspose.Cells | C# set tiled graphic on ODS worksheet page background | Assign background color and image to ODSPageBackground | Save ODS with custom page background using Aspose.Cells
// Developer Intent: The developer wants to programmatically apply a graphic (e.g., PNG) as a tiled page background, optionally add a solid background color, and export the workbook to ODS format using Aspose.Cells for .NET.
// Use Cases: Branding reports with a repeated logo behind data in ODS spreadsheets. | Creating ODS templates that include a light-colored watermark for visual hierarchy. | Generating financial statements where a corporate color scheme is applied to every page.
// AI Prompts: Write C# code with Aspose.Cells to set a JPEG image as a stretched, centered background for an ODS worksheet and save the file. | Show how to change the ODS page background color based on a cell value and apply a tiled PNG graphic. | Explain how to load a background image from a memory stream instead of a file and assign it to OdsPageBackground.

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Ods;

// Demonstrates how to create a workbook, access the ODSPageBackground of the first worksheet, assign a PNG image as a tiled graphic, set a background color, and save the file as an ODS document with Aspose.Cells for .NET.
class OdsPageBackgroundGraphicDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Access the ODS page background object
        OdsPageBackground background = sheet.PageSetup.ODSPageBackground;

        // Set the background type to graphic
        background.Type = OdsPageBackgroundType.Graphic;

        // Load an image file into a byte array and assign it as the graphic data
        string imagePath = "background.png"; // Replace with your image file path
        if (File.Exists(imagePath))
        {
            background.GraphicData = File.ReadAllBytes(imagePath);
        }

        // Define how the graphic should be displayed
        background.GraphicType = OdsPageBackgroundGraphicType.Tile;
        background.GraphicPositionType = OdsPageBackgroundGraphicPositionType.CenterCenter;

        // Optionally set a background color that will appear behind the graphic
        background.Color = Color.LightYellow;

        // Save the workbook as an ODS file
        workbook.Save("WorkbookWithGraphicBackground.ods");
    }
}
