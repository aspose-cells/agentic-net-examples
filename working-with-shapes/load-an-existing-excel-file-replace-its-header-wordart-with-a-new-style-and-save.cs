using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "OutputWorkbook.xlsx";

            try
            {
                // Load existing workbook or create a new one if the file is missing
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                    workbook.Save(inputPath); // optional placeholder
                }

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Get the collection of shapes on the worksheet
                ShapeCollection shapes = worksheet.Shapes;

                // Add a WordArt shape that will serve as the new header
                Shape header = shapes.AddWordArt(
                    PresetWordArtStyle.WordArtStyle5, // preset style
                    "New Header Title",               // text
                    0, 0,                             // top row & vertical offset (pixels)
                    0, 0,                             // left column & horizontal offset (pixels)
                    50,                               // height (pixels)
                    400);                             // width (pixels)

                // Apply additional formatting if the shape is WordArt
                if (header.IsWordArt)
                {
                    TextEffectFormat textEffect = header.TextEffect;
                    textEffect.FontBold = true;
                    textEffect.FontItalic = false;
                    textEffect.FontName = "Arial";
                    textEffect.FontSize = 24;
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}