// Title: Extract ODS Page Background Color Using Aspose.Cells for .NET
// Description: Loads an ODS workbook, accesses the first worksheet's PageSetup, reads the OdsPageBackground object, determines if it is a solid color, outputs the color value for audit, and saves the file unchanged.
// Keywords: Aspose.Cells OdsPageBackground | ODS page background color .NET | read ODS page setup background | C# ODS background audit | Aspose.Cells OdsPageBackgroundType
// Common Searches: how to get ODS page background color with Aspose.Cells | Aspose.Cells retrieve ODS background type C# | audit ODS workbook page background color | C# read ODS page setup background
// Developer Intent: Obtain and log the solid‑color background of an ODS worksheet for compliance or reporting.
// Use Cases: Verify that an ODS sheet uses a specific solid color before publishing. | Detect whether the page background is a graphic, color, or none for conditional processing. | Perform a non‑destructive audit of ODS files by reading and recording background settings.
// AI Prompts: Generate C# code with Aspose.Cells that returns the OdsPageBackgroundType and, if it is Color, the hex value of the background. | Create a function that logs the ODS page background color or indicates if a graphic is used. | Write a unit test that confirms the background color extraction works for an ODS file containing a solid color.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Ods;

namespace OdsPageBackgroundAudit
{
    // Loads an ODS workbook, accesses the first worksheet's PageSetup, reads the OdsPageBackground object, determines if it is a solid color, outputs the color value for audit, and saves the file unchanged.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the ODS file to audit
            string inputPath = "input.ods";

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the PageSetup object
            PageSetup pageSetup = sheet.PageSetup;

            // Retrieve the ODS page background object
            OdsPageBackground background = pageSetup.ODSPageBackground;

            // Check if the background type is set to Color
            if (background.Type == OdsPageBackgroundType.Color)
            {
                // Get the background color
                Color bgColor = background.Color;

                // Output the color information for auditing
                Console.WriteLine($"ODS Page Background Type: Color");
                Console.WriteLine($"Background Color: {bgColor}");
            }
            else
            {
                // If not a color background, indicate the current type
                Console.WriteLine($"ODS Page Background Type: {background.Type}");
                if (background.Type == OdsPageBackgroundType.Graphic)
                {
                    Console.WriteLine("Background is set to a graphic. Color information is not applicable.");
                }
                else
                {
                    Console.WriteLine("No background is set.");
                }
            }

            // Optionally save the workbook unchanged (lifecycle rule: save)
            // This demonstrates adherence to the save rule without altering the file.
            workbook.Save("output_audit.ods");
        }
    }
}
