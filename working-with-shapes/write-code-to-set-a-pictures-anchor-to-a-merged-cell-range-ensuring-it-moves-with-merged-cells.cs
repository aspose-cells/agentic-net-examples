// Title: C# – Anchor a Picture to a Merged Cell Range with MoveAndSize using Aspose.Cells
// Description: Demonstrates how to merge cells (B2:D4), add a picture, anchor it to the merged range with picture.MoveToRange, set PlacementType.MoveAndSize so the image moves and resizes with the cells, and save the workbook.
// Keywords: Aspose.Cells | C# | picture anchor merged cells | MoveToRange | PlacementType.MoveAndSize | add image to worksheet | Excel picture example | GitHub Aspose.Cells sample | code snippet
// Common Searches: Aspose.Cells anchor image to merged cells C# | picture.MoveToRange merged range example | set picture placement MoveAndSize Aspose.Cells | add picture to merged cell range .NET | Aspose.Cells picture moves with rows and columns
// Developer Intent: Place a picture inside a merged cell range and have it move and resize automatically with the cells.
// Use Cases: Insert a company logo into a merged header that stays aligned when the header size changes. | Embed a product thumbnail in a merged block that expands with added data rows. | Create a dashboard where charts are placed in merged cells and adjust automatically with layout modifications.
// AI Prompts: Generate C# code using Aspose.Cells to add a picture to merged cells B2:D4 and set its placement to MoveAndSize. | Show an Aspose.Cells .NET example that anchors an image to a merged range and updates its position when rows or columns are inserted. | Explain the effect of picture.MoveToRange and picture.Placement on image behavior in merged cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to merge cells (B2:D4), add a picture, anchor it to the merged range with picture.MoveToRange, set PlacementType.MoveAndSize so the image moves and resizes with the cells, and save the workbook.
    public class PictureInMergedCell
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Merge cells B2:D4 (rows 1-3, columns 1-3)
            worksheet.Cells.Merge(1, 1, 3, 3);

            // Path to the image file
            string imagePath = "sample.png";

            // Verify that the image file exists to avoid FileNotFoundException
            if (!File.Exists(imagePath))
            {
                throw new FileNotFoundException($"Image file not found: {imagePath}");
            }

            // Add a picture. The picture is initially placed at the top‑left cell of the range.
            int pictureIndex = worksheet.Pictures.Add(1, 1, imagePath);
            Picture picture = worksheet.Pictures[pictureIndex];

            // Anchor the picture to the merged range so it moves and resizes with the cells.
            picture.MoveToRange(1, 1, 3, 3);

            // Ensure the picture moves and sizes with the cells.
            picture.Placement = PlacementType.MoveAndSize;

            // Save the workbook
            string outputPath = "PictureInMergedCell.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
    }
}
