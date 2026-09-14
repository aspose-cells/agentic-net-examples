// Title: Extract graphic page background from an ODS workbook and save it as a PNG file using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an ODS file with Aspose.Cells, accesses the first worksheet's ODSPageBackground, verifies the background type is Graphic, and saves the embedded image bytes to a PNG file. | Show how to check for a graphic ODSPageBackground in a workbook and export the background image to disk as PNG using the Aspose.Cells .NET API.
// Common Searches: C# Aspose.Cells extract ODS worksheet page background image to PNG | How to retrieve embedded graphic background from an ODS file using Aspose.Cells .NET | Save ODS page background as PNG with Aspose.Cells in a console application | Aspose.Cells OdsPageBackgroundType Graphic example for extracting background image | Extract ODS file background picture and write to file using Aspose.Cells for .NET
// Tags: extract ODS page background graphic Aspose.Cells | save ODS background image as PNG C# | OdsPageBackground graphic data extraction .NET | Aspose.Cells export embedded ODS background | write ODS worksheet background to file

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Ods;

// The example loads an ODS workbook with Aspose.Cells, accesses the first worksheet's ODSPageBackground, confirms it is a graphic type, and writes the embedded PNG bytes to a specified output file.
class ExtractOdsBackground
{
    static void Main()
    {
        // Path to the source ODS file
        string odsFilePath = "input.ods";

        // Path where the extracted PNG will be saved
        string pngOutputPath = "background.png";

        // Load the ODS workbook
        Workbook workbook = new Workbook(odsFilePath);

        // Get the ODS page background of the first worksheet
        OdsPageBackground background = workbook.Worksheets[0].PageSetup.ODSPageBackground;

        // Check that the background is of graphic type and contains data
        if (background.Type == OdsPageBackgroundType.Graphic && background.GraphicData != null && background.GraphicData.Length > 0)
        {
            // Write the graphic data (assumed to be PNG) to the output file
            File.WriteAllBytes(pngOutputPath, background.GraphicData);
            Console.WriteLine($"Graphic background extracted successfully to: {pngOutputPath}");
        }
        else
        {
            Console.WriteLine("The ODS file does not contain a graphic background.");
        }
    }
}
