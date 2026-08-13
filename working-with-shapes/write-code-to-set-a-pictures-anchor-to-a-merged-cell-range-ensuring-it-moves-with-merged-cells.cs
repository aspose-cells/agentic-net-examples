// Title: C# – Anchor an Image to a Merged Cell Range and Enable MoveAndSize with Aspise.Cells
// Description: The sample builds a workbook, merges cells B2:D4, adds a PNG picture, marks it as placed in a cell, sets its placement to MoveAndSize, positions it over the merged block via picture.MoveToRange, and writes the result to PictureAnchoredToMergedCell.xlsx.
// Keywords: Aspose.Cells C# picture anchor | merged cells image placement | MoveAndSize placement type | IsPlacedInCell property | picture.MoveToRange example | Excel picture positioning programmatically | Aspose.Cells picture handling | C# Excel image merge | dynamic image alignment in Excel | Aspose.Cells API picture
// Common Searches: Aspose.Cells anchor image to merged cells C# | how to make picture move and resize with merged range in .NET | set picture placement MoveAndSize Aspose.Cells | C# picture.MoveToRange merged block example | IsPlacedInCell usage Aspose.Cells
// Developer Intent: Insert an image so it stays aligned with a merged block and automatically moves or resizes when the underlying rows or columns change.
// Use Cases: Add a company logo to a merged header that expands with column width adjustments. | Place a snapshot of a chart inside a merged reporting area, preserving layout after data updates. | Embed a watermark that remains correctly positioned when users edit row heights or column widths.
// AI Prompts: Generate C# code using Aspose.Cells to anchor a PNG to the merged range B2:D4 and enable MoveAndSize behavior. | Explain the impact of IsPlacedInCell and Placement properties on picture dynamics in an Excel worksheet. | Show how to programmatically confirm that a picture is anchored to a merged area after saving the workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // The sample builds a workbook, merges cells B2:D4, adds a PNG picture, marks it as placed in a cell, sets its placement to MoveAndSize, positions it over the merged block via picture.MoveToRange, and writes the result to PictureAnchoredToMergedCell.xlsx.
    public class Program
    {
        public static void Main()
        {
            try
            {
                PictureAnchorToMergedCell.Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class PictureAnchorToMergedCell
    {
        public static void Run()
        {
            // Verify that the image file exists to avoid FileNotFoundException
            const string imagePath = "sample.png";
            if (!File.Exists(imagePath))
                throw new FileNotFoundException($"Image file not found: {imagePath}");

            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];

            // Define and merge the cell range B2:D4 (rows 1‑3, columns 1‑3 zero‑based)
            worksheet.Cells.Merge(1, 1, 3, 3);

            // Add the picture to the worksheet; initial position is top‑left cell (A1)
            int pictureIndex = worksheet.Pictures.Add(0, 0, imagePath);
            var picture = worksheet.Pictures[pictureIndex];

            // Anchor the picture to cells and make it move/size with the merged range
            picture.IsPlacedInCell = true;                     // anchor to cells
            picture.Placement = PlacementType.MoveAndSize;    // move and resize with cells

            // Anchor the picture to the merged cell range B2:D4
            // MoveToRange(topRow, leftColumn, bottomRow, rightColumn) – zero‑based indices
            picture.MoveToRange(1, 1, 3, 3);

            // Save the workbook
            const string outputPath = "PictureAnchoredToMergedCell.xlsx";
            workbook.Save(outputPath);
        }
    }
}
