// Title: Set a Solid Blue ODS Page Background with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, access the first worksheet, configure the OdsPageBackground to a solid blue color, and save the result as an ODS file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | ODS page background | solid color background | OdsPageBackground | Color.Blue | save ODS file | OpenDocument Spreadsheet | programmatic ODS styling | Excel to ODS conversion
// Common Searches: Aspose.Cells set ODS page background color C# | how to apply solid color background to ODS with Aspose | OdsPageBackground blue example .NET | save workbook as ODS with custom background | programmatically change ODS page background
// Developer Intent: Apply a solid blue background to an ODS worksheet and save the file using Aspose.Cells for .NET.
// Use Cases: Generate branded ODS reports by applying a corporate color to the page background before export. | Create a template workbook where the first sheet automatically uses a blue background for visual emphasis. | Update existing ODS documents to match a new visual style by programmatically changing the page background color.
// AI Prompts: Write C# code with Aspose.Cells that sets the ODS page background to a custom RGB value and saves the workbook as an ODS file. | Explain the relationship between OdsPageBackground.Type and OdsPageBackground.Color when applying a solid color background in an ODS document. | Provide a sample that iterates through all worksheets in a workbook and assigns each a different solid background color before saving as ODS.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Ods;

namespace OdsPageBackgroundExample
{
    // Demonstrates how to create a workbook, access the first worksheet, configure the OdsPageBackground to a solid blue color, and save the result as an ODS file using Aspose.Cells for .NET.
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

            // Set the background type to Color and apply a solid blue color
            background.Type = OdsPageBackgroundType.Color;
            background.Color = Color.Blue;

            // Save the workbook as an ODS file with the background applied
            workbook.Save("OdsPageBackgroundBlue.ods");
        }
    }
}
