using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtWatermark
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add sample data (optional, just to have some content)
                sheet.Cells["A1"].PutValue("Sample Data");
                sheet.Cells["B2"].PutValue(123);

                // Add a WordArt (text effect) shape that will act as a watermark
                // Parameters: preset effect, text, font name, font size, bold, italic,
                // start row, start column, end row, end column, width, height
                Shape watermark = sheet.Shapes.AddTextEffect(
                    MsoPresetTextEffect.TextEffect1,   // preset effect
                    "CONFIDENTIAL",                    // watermark text
                    "Arial",                           // font name
                    72,                                // font size
                    true,                              // bold
                    false,                             // italic
                    5, 0, 15, 5,                       // position (rows/columns)
                    300, 100);                         // width and height of the shape

                // Set shape properties to make it look like a typical watermark
                watermark.RotationAngle = -45;                 // diagonal orientation
                watermark.Fill.Transparency = 0.8;             // semi‑transparent
                watermark.Line.Weight = 0;                     // hide border
                watermark.ZOrderPosition = 0;                  // send to back

                // Save the workbook as XLSX preserving all shape attributes
                workbook.Save("WordArtWatermark.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}