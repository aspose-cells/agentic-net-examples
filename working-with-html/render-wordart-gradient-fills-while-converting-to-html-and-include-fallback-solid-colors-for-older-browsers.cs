using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace WordArtGradientHtmlDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook
                Workbook workbook = new Workbook();

                // 2. Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // 3. Add a WordArt shape with a preset gradient style
                //    Parameters: style, text, upperLeftRow, top, upperLeftColumn, left, height, width
                Shape wordArt = sheet.Shapes.AddWordArt(
                    PresetWordArtStyle.WordArtStyle7,
                    "Gradient WordArt",
                    2, 0,   // upper left row, top offset
                    2, 0,   // upper left column, left offset
                    100,    // height (points)
                    400);   // width (points)

                // 4. Prepare HTML save options with WordArt exported as images (fallback for older browsers)
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    // ExportWordArtAsImage is not available in this version; WordArt will be rendered as part of the HTML.
                    ExportActiveWorksheetOnly = true,    // Export only the active sheet
                    ExportImagesAsBase64 = true          // Embed images as Base64 strings
                };

                // 5. Define output file path
                string outputPath = "WordArtGradient.html";

                // Ensure the output directory exists (handle cases where outputPath has no directory component)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // 6. Save the workbook as HTML
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine("HTML file with gradient WordArt generated successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}