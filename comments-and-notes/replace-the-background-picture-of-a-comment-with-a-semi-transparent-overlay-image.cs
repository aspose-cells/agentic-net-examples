using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ReplaceCommentBackground
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a comment to cell B2
            int commentIdx = sheet.Comments.Add("B2");
            Comment comment = sheet.Comments[commentIdx];
            comment.Note = "Original comment";

            // Load overlay image if it exists
            string overlayPath = "overlay.png";
            if (!File.Exists(overlayPath))
                throw new FileNotFoundException($"Overlay image not found: {overlayPath}");

            byte[] overlayBytes = File.ReadAllBytes(overlayPath);

            // Set the overlay image as the comment background
            comment.CommentShape.Fill.ImageData = overlayBytes;   // picture fill
            comment.CommentShape.Fill.Transparency = 0.5;        // 50 % transparent

            // Save the workbook
            workbook.Save("CommentWithOverlay.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}