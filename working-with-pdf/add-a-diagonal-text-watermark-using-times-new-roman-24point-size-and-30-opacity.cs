// Title: Create a diagonal Times New Roman 24‑pt text watermark with 30% opacity in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Add a TextEffect shape with the text "Confidential" in Times New Roman 24 pt to the first worksheet, rotate it –45°, and set its fill transparency to 0.7 using Aspose.Cells C#. | Configure the shape’s font color and adjust its width and height to span the worksheet while maintaining a diagonal orientation in Aspose.Cells. | Save the modified workbook as WatermarkedWorkbook.xlsx after applying the semi‑transparent diagonal watermark with Aspose.Cells for .NET.
// Common Searches: how to add a diagonal text watermark to an Excel file with Aspose.Cells C# | Aspose.Cells set watermark transparency to 30 percent in .NET | rotate text effect shape -45 degrees Aspose.Cells example | create Times New Roman 24pt watermark in Excel using Aspose.Cells | save workbook after applying watermark Aspose.Cells C#
// Tags: Aspose.Cells TextEffect watermark implementation | Aspose.Cells shape transparency configuration | Aspose.Cells diagonal orientation of watermark shape | Aspose.Cells custom font usage in watermark | Aspose.Cells save workbook with watermark

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace WatermarkExample
{
    // Demonstrates creating a new workbook, adding a TextEffect shape containing the text "Confidential" in Times New Roman 24 pt, rotating it –45° to appear diagonal, setting fill transparency to 0.7 (30 % opacity), and saving the file as WatermarkedWorkbook.xlsx with Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add a diagonal text watermark using a text effect shape
                // Parameters: preset effect, text, font name, font size, bold, italic,
                // upper-left cell (row, column), lower-right cell (row, column), width, height
                Shape watermarkShape = sheet.Shapes.AddTextEffect(
                    MsoPresetTextEffect.TextEffect1,
                    "Confidential",
                    "Times New Roman",
                    24,
                    false,
                    false,
                    0,   // upper-left row
                    0,   // upper-left column
                    20,  // lower-right row
                    10,  // lower-right column
                    200, // width of the shape
                    50   // height of the shape
                );

                // Rotate the shape to create a diagonal appearance
                watermarkShape.RotationAngle = -45;

                // Set transparency (0 = fully opaque, 1 = fully transparent)
                // 0.7 makes the watermark 30% opaque
                watermarkShape.Fill.Transparency = 0.7;

                // Optionally set the text color
                watermarkShape.Font.Color = Color.Black;

                // Define output file path
                string outputPath = "WatermarkedWorkbook.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
