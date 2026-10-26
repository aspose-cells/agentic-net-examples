// Title: How to anchor a comment shape to cell G12 with MoveAndSize placement using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# attach comment box to a specific cell so it moves with the cell | set comment shape placement to MoveAndSize for cell G12 in Aspose.Cells | link Excel comment shape to a worksheet cell programmatically with Aspose.Cells | C# example of anchoring a comment shape to a cell and preserving size on row/column changes | Aspose.Cells comment shape dynamic positioning when rows or columns are inserted
// Tags: comment shape placement MoveAndSize | anchor comment shape to cell | dynamic comment positioning Aspose.Cells | set comment shape size programmatically | link comment shape to cell G12

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example creates a new workbook, adds a comment to cell G12, sets the author and note, configures the comment's shape width, height, and Placement to MoveAndSize so the shape follows the cell when it moves, and saves the file as LinkedCommentBox.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Target cell G12 (zero‑based indexes)
                int targetColumn = 6; // G
                int targetRow = 11;   // 12

                // Add a comment to the target cell (older API may require only row/column)
                int commentIndex = worksheet.Comments.Add(targetRow, targetColumn);
                Comment comment = worksheet.Comments[commentIndex];
                comment.Author = "Author";
                comment.Note = "This is a comment linked to G12.";

                // Configure the comment shape (size and placement)
                comment.CommentShape.Width = 200;   // width in points
                comment.CommentShape.Height = 100; // height in points
                comment.CommentShape.Placement = PlacementType.MoveAndSize;

                // Define output path
                string outputPath = "LinkedCommentBox.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
