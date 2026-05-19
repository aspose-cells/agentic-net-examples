using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsWatermarkDemo
{
    public class WatermarkHelper
    {
        /// <summary>
        /// Removes all WordArt shapes from every worksheet, adds a text watermark, and saves as PDF.
        /// </summary>
        public static void RemoveWordArtAndAddWatermark(string inputFile, string outputFile, string watermarkText)
        {
            try
            {
                if (!File.Exists(inputFile))
                    throw new FileNotFoundException($"Input file not found: {inputFile}");

                // Load the workbook.
                Workbook workbook = new Workbook(inputFile);

                // Remove WordArt shapes.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    ShapeCollection shapes = sheet.Shapes;
                    for (int i = shapes.Count - 1; i >= 0; i--)
                    {
                        Shape shape = shapes[i];
                        if (shape.IsWordArt)
                            shapes.RemoveAt(i);
                    }
                }

                // Create rendering font for the watermark.
                RenderingFont font = new RenderingFont("Arial", 48)
                {
                    Bold = true,
                    Color = Color.Red,
                    Italic = false
                };

                // Create the watermark.
                RenderingWatermark watermark = new RenderingWatermark(watermarkText, font)
                {
                    Rotation = 45f,
                    Opacity = 0.25f,
                    IsBackground = true,
                    HAlignment = TextAlignmentType.Center,
                    VAlignment = TextAlignmentType.Center,
                    ScaleToPagePercent = 80
                };

                // Set PDF save options with the watermark.
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    Watermark = watermark
                };

                // Save as PDF.
                workbook.Save(outputFile, pdfOptions);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Expected arguments: inputFile outputFile watermarkText
            if (args.Length < 3)
            {
                Console.WriteLine("Usage: AsposeCellsWatermarkDemo <input.xlsx> <output.pdf> <watermark text>");
                return;
            }

            string inputFile = args[0];
            string outputFile = args[1];
            string watermarkText = args[2];

            WatermarkHelper.RemoveWordArtAndAddWatermark(inputFile, outputFile, watermarkText);
        }
    }
}