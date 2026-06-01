using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtExample
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Get the shapes collection of the worksheet
                ShapeCollection shapes = worksheet.Shapes;

                // Add a WordArt shape (preset style, text, initial position and size)
                Shape wordArt = shapes.AddWordArt(
                    PresetWordArtStyle.WordArtStyle1, // preset style
                    "Hello World",                    // text
                    0,    // top row index (initial, will be repositioned)
                    0,    // vertical offset in pixels
                    0,    // left column index (initial)
                    0,    // horizontal offset in pixels
                    100,  // height in pixels
                    300   // width in pixels
                );

                // Anchor the shape to a specific cell range (e.g., B3:D5)
                // Rows and columns are zero‑based indices
                int startRow = 2;    // B3 -> row index 2
                int startColumn = 1; // column B -> index 1
                int endRow = 4;      // D5 -> row index 4
                int endColumn = 3;   // column D -> index 3
                wordArt.MoveToRange(startRow, startColumn, endRow, endColumn);

                // Set the placement so the shape moves and resizes with the cells
                wordArt.Placement = PlacementType.MoveAndSize;

                // Define output file path
                string outputPath = "WordArtAnchored.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}