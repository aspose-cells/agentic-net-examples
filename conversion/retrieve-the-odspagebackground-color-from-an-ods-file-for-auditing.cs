using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Ods;

class Program
{
    static void Main()
    {
        // Load the existing ODS file
        Workbook workbook = new Workbook("input.ods");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Get the ODS page background object
        OdsPageBackground background = sheet.PageSetup.ODSPageBackground;

        // Retrieve the background type and color
        OdsPageBackgroundType bgType = background.Type;
        Color bgColor = background.Color;

        // Output the retrieved information for auditing
        Console.WriteLine("ODS Page Background Type: " + bgType);
        Console.WriteLine("ODS Page Background Color: " + bgColor);

        // Save the workbook (unchanged) if needed for further processing
        workbook.Save("output_audit.ods");
    }
}