// Title: Resize a WordArt shape to fill the entire worksheet with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set WordArt shape dimensions to cover whole sheet | How to anchor WordArt from first to last cell using Aspose.Cells .NET | Programmatically stretch WordArt to match worksheet size in Aspose.Cells | Apply MoveAndSize placement to WordArt and span all rows and columns Aspose.Cells | Calculate pixel width and height for WordArt shape in Aspose.Cells C#
// Tags: WordArt shape resizing Aspose.Cells | cell‑based anchoring UpperLeftRow LowerRightRow Aspose.Cells | pixel‑based sizing Top Left Width Height Aspose.Cells | shape placement MoveAndSize Aspose.Cells | calculate total column width pixels Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace WordArtFullWorksheet
{
    // The example loads or creates a workbook, adds a WordArt shape, and expands it to cover the entire worksheet by setting UpperLeftRow/Column and LowerRightRow/Column. It also sets the shape's Placement to MoveAndSize, optionally computes the total pixel width and height of the target range, applies Top, Left, Width, and Height, and saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the input workbook (if it does not exist we will create a new one)
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            Workbook workbook = null;

            try
            {
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                    Console.WriteLine($"Loaded workbook from \"{inputPath}\".");
                }
                else
                {
                    workbook = new Workbook();
                    Console.WriteLine($"Input file not found. Created a new workbook.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load or create workbook: {ex.Message}");
                return;
            }

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one cell so that the worksheet has a defined grid
            worksheet.Cells["A1"].PutValue("Placeholder");

            // Add a WordArt shape with minimal initial size (it will be resized later)
            // Signature: AddWordArt(PresetWordArtStyle style, string text,
            //                         int topRow, int top, int leftColumn, int left,
            //                         int height, int width)
            Shape wordArt = worksheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1,
                "Full Worksheet WordArt",
                0,      // topRow
                0,      // top offset (pixels)
                0,      // leftColumn
                0,      // left offset (pixels)
                10,     // height (pixels) – temporary
                10);    // width (pixels) – temporary

            Console.WriteLine("WordArt shape added.");

            // ------------------------------------------------------------
            // Option 1 – Use cell‑based anchoring (UpperLeft/LowerRight)
            // ------------------------------------------------------------
            // Determine the furthest occupied row and column.
            // If you want to cover the *entire* sheet regardless of data,
            // you could use the maximum possible rows/columns (e.g., 0‑65535 and 0‑255).
            // Here we use the last used row/column to keep the file size reasonable.
            int maxRow = worksheet.Cells.MaxDataRow;      // zero‑based index of last row with data
            int maxCol = worksheet.Cells.MaxDataColumn;   // zero‑based index of last column with data

            // If the sheet is empty, fall back to a default range (e.g., rows 0‑20, cols 0‑10)
            if (maxRow < 0) maxRow = 20;
            if (maxCol < 0) maxCol = 10;

            // Stretch the WordArt to cover the range from (0,0) to (maxRow, maxCol)
            wordArt.UpperLeftRow = 0;
            wordArt.UpperLeftColumn = 0;
            wordArt.LowerRightRow = maxRow;
            wordArt.LowerRightColumn = maxCol;

            // Ensure the shape moves and resizes with the cells
            wordArt.Placement = PlacementType.MoveAndSize;

            Console.WriteLine($"WordArt anchored from row 0, column 0 to row {maxRow}, column {maxCol} (cell‑based).");

            // ------------------------------------------------------------
            // Option 2 – Use pixel‑based dimensions (Top/Left/Width/Height)
            // ------------------------------------------------------------
            // Calculate total pixel width and height of the selected range.
            // This is optional; you can comment out this block if you prefer
            // the cell‑based anchoring above.
            try
            {
                double totalWidthPixels = 0;
                for (int col = 0; col <= maxCol; col++)
                {
                    // GetColumnWidthPixel returns the width of a column in pixels.
                    totalWidthPixels += worksheet.Cells.GetColumnWidthPixel(col);
                }

                double totalHeightPixels = 0;
                for (int row = 0; row <= maxRow; row++)
                {
                    // GetRowHeightPixel returns the height of a row in pixels.
                    totalHeightPixels += worksheet.Cells.GetRowHeightPixel(row);
                }

                // Apply pixel‑based positioning
                wordArt.Top = 0;
                wordArt.Left = 0;
                wordArt.Width = (int)Math.Ceiling(totalWidthPixels);
                wordArt.Height = (int)Math.Ceiling(totalHeightPixels);

                Console.WriteLine($"WordArt resized to {wordArt.Width}x{wordArt.Height} pixels (pixel‑based).");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Pixel‑based sizing failed: {ex.Message}");
                // Continue – the cell‑based anchoring is already applied.
            }

            // ------------------------------------------------------------
            // Save the workbook
            // ------------------------------------------------------------
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
