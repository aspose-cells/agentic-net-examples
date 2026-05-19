using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class WordArtTransparencyDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a WordArt shape with a preset style
                Shape wordArt = worksheet.Shapes.AddWordArt(
                    PresetWordArtStyle.WordArtStyle2,
                    "Background WordArt",
                    2,    // top row index
                    0,    // vertical offset in pixels
                    2,    // left column index
                    0,    // horizontal offset in pixels
                    100,  // height in pixels
                    400   // width in pixels
                );

                // Set the fill transparency to 30%
                wordArt.Fill.Transparency = 0.3;

                // Hide the shape outline for a cleaner background effect
                // (LineFormat.IsVisible is not available in some versions; set weight to 0 instead)
                wordArt.Line.Weight = 0;

                // Save the workbook
                string outputPath = "WordArtTransparencyDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            WordArtTransparencyDemo.Run();
        }
    }
}