using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

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

            // Add a WordArt shape with a gradient fill (preset style 7)
            // Parameters: style, text, upperLeftRow, top, upperLeftColumn, left, height, width
            Shape wordArt = sheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle7,
                "Aspose.Cells WordArt",
                2, 0,   // row, top offset
                2, 0,   // column, left offset
                100,    // height
                400);   // width

            // (SetPosition is not required here; AddWordArt already positions the shape)

            // Configure HTML save options for standards‑compliant output
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                HtmlVersion = HtmlVersion.Html5,
                ExportWorksheetCSSSeparately = true,
                ExcludeUnusedStyles = false,
                DisableCss = false
            };

            // Define output path
            string outputPath = "WordArtGradient.html";

            // Ensure the output directory exists (handle possible null directory)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as HTML
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook with WordArt saved to HTML at: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}