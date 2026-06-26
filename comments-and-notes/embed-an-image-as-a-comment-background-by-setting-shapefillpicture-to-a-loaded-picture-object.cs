using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

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
            int commentIndex = sheet.Comments.Add(2, 2);
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "Comment with image background";

            // Load image bytes if the file exists
            string imagePath = "background.png";
            if (File.Exists(imagePath))
            {
                byte[] imageData = File.ReadAllBytes(imagePath);
                // Set the comment shape's fill to use the loaded image
                comment.CommentShape.Fill.ImageData = imageData;
            }
            else
            {
                Console.WriteLine($"Image file not found: {imagePath}. The comment will be saved without a background image.");
            }

            // Optionally adjust the comment shape size to better display the image
            comment.CommentShape.Width = 200;
            comment.CommentShape.Height = 150;

            // Save the workbook
            workbook.Save("CommentWithImageBackground.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}