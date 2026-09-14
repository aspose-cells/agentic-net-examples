// Title: How to set a solid blue ODS page background using Aspose.Cells for .NET and save the workbook
// AI Prompts: Write C# code that creates a new Workbook, sets the ODS page background to a solid blue color via OdsPageBackground, and saves it as an .ods file. | Demonstrate using Aspose.Cells PageSetup.ODSPageBackground to apply a solid color background to the first worksheet in a .NET application. | Provide a complete example that configures OdsPageBackground.Type and OdsPageBackground.Color to produce a blue background in an ODS document.
// Common Searches: Aspose.Cells C# set ODS page background color to blue | How to apply solid color background to ODS file with Aspose.Cells .NET | Saving ODS workbook with custom page background using Aspose.Cells API | C# example for OdsPageBackground Type Color property | Change ODS page background to solid color in Aspose.Cells
// Tags: Aspose.Cells OdsPageBackground set solid color | C# ODS page background color | Aspose.Cells save ODS with background | PageSetup ODSPageBackground API | Workbook ODS blue background

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Ods;

namespace OdsPageBackgroundExample
{
    // // Creates a new Workbook, accesses the first worksheet, configures the ODSPageBackground to a solid blue color via PageSetup, and saves the file as OdsPageBackgroundBlue.ods.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the ODS page background object from the worksheet's page setup
            OdsPageBackground background = sheet.PageSetup.ODSPageBackground;

            // Set the background type to solid color
            background.Type = OdsPageBackgroundType.Color;

            // Apply a solid blue color as the background
            background.Color = Color.Blue;

            // Save the workbook as an ODS file with the background applied
            workbook.Save("OdsPageBackgroundBlue.ods");
        }
    }
}
