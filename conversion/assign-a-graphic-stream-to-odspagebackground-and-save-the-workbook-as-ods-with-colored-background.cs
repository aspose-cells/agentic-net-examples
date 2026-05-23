using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Ods;

namespace OdsBackgroundDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Access the ODS page background through PageSetup
            OdsPageBackground background = sheet.PageSetup.ODSPageBackground;

            // Set a solid background color
            background.Type = OdsPageBackgroundType.Color;
            background.Color = Color.LightGreen;

            // Change the type to graphic to assign an image stream
            background.Type = OdsPageBackgroundType.Graphic;

            // Load image data into a byte array (replace with a valid image path)
            string imagePath = "background.png";
            if (File.Exists(imagePath))
            {
                background.GraphicData = File.ReadAllBytes(imagePath);
                background.GraphicType = OdsPageBackgroundGraphicType.Tile;
                background.GraphicPositionType = OdsPageBackgroundGraphicPositionType.CenterCenter;
            }
            else
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Save the workbook as ODS
            workbook.Save("WorkbookWithGraphicBackground.ods");
            Console.WriteLine("Workbook saved successfully.");
        }
    }
}