// Title: Set a PNG image as the background of an Excel comment using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a PNG file into a byte array and assigns it to the comment shape's fill image using Aspose.Cells. | Show how to resize the comment shape after applying an image background and then save the workbook as an .xlsx file with Aspose.Cells.
// Common Searches: Aspose.Cells C# set image background for a cell comment | how to use CommentShape.Fill.ImageData to embed a picture in an Excel comment | load PNG bytes and apply as comment fill in Aspose.Cells .NET | resize Excel comment shape after adding image background with Aspose.Cells | save workbook with image‑filled comment using Aspose.Cells C#
// Tags: comment shape image fill Aspose.Cells | png background for Excel comment .NET | resize comment shape dimensions C# | load image bytes for comment fill Aspose.Cells | save workbook with image comment Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a comment to cell C3, reads a PNG file into a byte array, assigns the image data to the comment's shape fill, optionally resizes the comment shape, and saves the workbook as CommentWithImage.xlsx.
class EmbedImageInComment
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a comment to cell C3 (row index 2, column index 2)
            int commentIdx = sheet.Comments.Add(2, 2);
            Comment comment = sheet.Comments[commentIdx];
            comment.Note = "Comment with image background";

            // Load the image file into a byte array if it exists
            string imagePath = "background.png"; // replace with your image file
            if (File.Exists(imagePath))
            {
                byte[] imageData = File.ReadAllBytes(imagePath);
                // Set the comment shape's fill to use the loaded image
                comment.CommentShape.Fill.ImageData = imageData;
            }
            else
            {
                Console.WriteLine($"Image file not found: {imagePath}. Skipping image background.");
            }

            // Optionally resize the comment shape
            comment.CommentShape.Width = 250;
            comment.CommentShape.Height = 150;

            // Save the workbook
            string outputPath = "CommentWithImage.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
