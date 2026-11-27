// Title: How to anchor a picture to a merged cell range and enable move‑and‑size behavior with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# anchor image to merged cells B2:D4 | set picture PlacementType.MoveAndSize for merged range in Aspose.Cells | C# example for setting picture UpperLeftRow and LowerRightColumn in Aspose.Cells | keep picture aligned with merged cells when rows or columns are resized Aspose.Cells | add picture that spans multiple cells without using IsPlacedInCell in Aspose.Cells
// Tags: Aspose.Cells picture anchoring to merged cells | C# set picture bounds UpperLeftRow LowerRightRow | PlacementType.MoveAndSize worksheet image | Add picture spanning merged range Aspose.Cells | IsPlacedInCell false multi‑cell picture

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace PictureAnchorExample
{
    // The example creates a workbook, merges cells B2:D4, inserts a tiny PNG picture, sets the picture's UpperLeftRow/Column and LowerRightRow/Column to cover the merged range, applies Placement = MoveAndSize and IsPlacedInCell = false so the image moves and resizes with the cells, adds a blue border for visibility, and saves the file as PictureAnchoredToMergedCells.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Define a merged cell range B2:D4 (zero‑based indices: rows 1‑3, columns 1‑3)
                int firstRow = 1;      // B2 row index
                int firstColumn = 1;   // B2 column index
                int lastRow = 3;       // D4 row index
                int lastColumn = 3;    // D4 column index

                // Merge the cells
                sheet.Cells.Merge(firstRow, firstColumn, lastRow - firstRow + 1, lastColumn - firstColumn + 1);
                // Optional: put some text in the merged area to visualize it
                sheet.Cells[firstRow, firstColumn].PutValue("Merged Area");

                // Prepare a tiny PNG image (1x1 red pixel) from a Base64 string
                string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";
                byte[] imageBytes = Convert.FromBase64String(base64Png);
                using (MemoryStream imgStream = new MemoryStream(imageBytes))
                {
                    // Add the picture. The Add method requires the upper‑left cell of the picture.
                    // We'll initially place it at the same upper‑left cell of the merged range.
                    int pictureIndex = sheet.Pictures.Add(firstRow, firstColumn, imgStream);
                    Picture picture = sheet.Pictures[pictureIndex];

                    // Anchor the picture to the merged cell range
                    picture.UpperLeftRow = firstRow;
                    picture.UpperLeftColumn = firstColumn;
                    picture.LowerRightRow = lastRow;
                    picture.LowerRightColumn = lastColumn;

                    // Ensure the picture moves and resizes together with the cells
                    picture.Placement = PlacementType.MoveAndSize;

                    // Optionally, you can control whether the picture is placed inside a single cell.
                    // For an anchored picture spanning multiple cells, keep IsPlacedInCell = false (default).
                    picture.IsPlacedInCell = false;

                    // Set a visible border so the picture area can be seen in the output file
                    picture.BorderLineColor = System.Drawing.Color.Blue;
                    picture.BorderWeight = 2;
                }

                // Save the workbook
                string outputPath = "PictureAnchoredToMergedCells.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
