using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class WordArtBackgroundDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add sample content to demonstrate the watermark effect
                sheet.Cells["A1"].PutValue("This is a sample worksheet with a background WordArt watermark.");

                // Insert a WordArt shape
                // Parameters: style, text, topRow, top, leftColumn, left, height, width
                Shape wordArt = sheet.Shapes.AddWordArt(
                    PresetWordArtStyle.WordArtStyle2, // preset style
                    "CONFIDENTIAL",                  // watermark text
                    5, 0,                            // top row and vertical offset (pixels)
                    2, 0,                            // left column and horizontal offset (pixels)
                    200, 600);                       // height and width (pixels)

                // Send the WordArt shape to the back so it acts as a background watermark
                // Passing a negative value moves the shape behind other objects
                wordArt.ToFrontOrBack(-1);

                // Save the workbook
                workbook.Save("WordArtBackgroundDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point required for console application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}