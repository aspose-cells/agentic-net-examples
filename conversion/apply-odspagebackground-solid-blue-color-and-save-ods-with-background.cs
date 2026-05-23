using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Ods;

namespace OdsPageBackgroundExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the ODS page background object
            OdsPageBackground background = sheet.PageSetup.ODSPageBackground;

            // Set background type to solid color
            background.Type = OdsPageBackgroundType.Color;

            // Apply solid blue color
            background.Color = Color.Blue;

            // Save the workbook as ODS with the background applied
            workbook.Save("OdsPageBackgroundBlue.ods");
        }
    }
}