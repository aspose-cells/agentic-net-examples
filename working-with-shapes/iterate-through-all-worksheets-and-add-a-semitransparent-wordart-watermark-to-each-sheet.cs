using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class AddWordArtWatermarkToAllSheets
{
    static void Main()
    {
        try
        {
            // Create a new workbook (contains a default worksheet)
            Workbook workbook = new Workbook();

            // Ensure worksheet names are unique – rename the default sheet and add new ones
            workbook.Worksheets[0].Name = "DefaultSheet";

            workbook.Worksheets.Add("Sheet1");
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Iterate through each worksheet and add a WordArt watermark
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];
                ShapeCollection shapes = sheet.Shapes;

                // Add WordArt shape as watermark
                Shape wordArt = shapes.AddWordArt(
                    PresetWordArtStyle.WordArtStyle1, // preset style
                    "CONFIDENTIAL",                   // watermark text
                    0, 0,                             // top row & vertical offset (pixels)
                    0, 0,                             // left column & horizontal offset (pixels)
                    100,                              // height (pixels)
                    400);                             // width (pixels)

                // Make the WordArt semi‑transparent
                wordArt.FillFormat.Transparency = 0.5; // 0 = opaque, 1 = fully transparent

                // Hide the outline
                wordArt.LineFormat.IsVisible = false;

                // Send the shape to the background
                wordArt.ZOrderPosition = -1;
            }

            // Define output file path
            string outputPath = "WorkbookWithWordArtWatermark.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}