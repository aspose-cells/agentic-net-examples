using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class WordArtFlipDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a WordArt shape with a preset style
                // Parameters: style, text, upper left row, upper left column,
                // row offset, column offset, height (pixels), width (pixels)
                Shape wordArt = worksheet.Shapes.AddWordArt(
                    PresetWordArtStyle.WordArtStyle1,
                    "Flip Me",
                    2,   // upper left row index
                    2,   // upper left column index
                    0,   // row offset (pixels)
                    0,   // column offset (pixels)
                    100, // height (pixels)
                    300  // width (pixels)
                );

                // Verify that the shape is recognized as WordArt
                Console.WriteLine("IsWordArt: " + wordArt.IsWordArt);

                // Output the initial flip state
                Console.WriteLine("Initial IsFlippedHorizontally: " + wordArt.IsFlippedHorizontally);

                // Flip the shape horizontally
                wordArt.IsFlippedHorizontally = true;

                // Output the new flip state to confirm the change
                Console.WriteLine("After flip IsFlippedHorizontally: " + wordArt.IsFlippedHorizontally);

                // Save the workbook to visualize the flipped WordArt
                string outputPath = "WordArtFlipped_Output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved to: " + outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            WordArtFlipDemo.Run();
        }
    }
}