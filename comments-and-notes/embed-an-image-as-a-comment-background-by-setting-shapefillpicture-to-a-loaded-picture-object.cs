// Title: Set an Image as the Background of a Cell Comment using Aspose.Cells for .NET (C#)
// Description: The sample creates a new workbook, adds a comment to cell C3, loads a PNG file (background.png) into a byte array, assigns the image to the comment's shape fill as a texture, and saves the workbook as CommentWithImage.xlsx.
// Keywords: Aspose.Cells comment background image | C# comment shape fill texture | Excel comment image fill Aspose | load PNG for comment background .NET | set comment shape picture Aspose.Cells | embed picture in Excel comment C# | Aspose.Cells FillType.Texture | comment shape Fill.ImageData
// Common Searches: How to add a picture to an Excel comment using Aspose.Cells C# | Aspose.Cells set comment background image | C# load image and apply to comment shape | Aspose.Cells Fill.ImageData example | Excel comment with image background .NET
// Developer Intent: Apply a PNG or JPEG as the background of a worksheet comment.
// Use Cases: Brand a comment with a company logo for internal documentation. | Show product thumbnail inside a comment to aid sales reports. | Provide a visual watermark in comments for data validation. | Create instructional notes with a diagram background for training materials. | Add a map snapshot as a comment background for location‑based data.
// AI Prompts: Generate C# code that sets a JPEG file as the background of a cell comment using Aspose.Cells. | Explain how to use a memory stream to assign a base64‑encoded image to a comment's fill in Aspose.Cells. | Show how to adjust the opacity of a comment background image with Aspose.Cells for .NET. | Provide steps to replace an existing comment background with a new picture programmatically. | Demonstrate how to retrieve and modify the Fill.Type of a comment shape in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, adds a comment to cell C3, loads a PNG file (background.png) into a byte array, assigns the image to the comment's shape fill as a texture, and saves the workbook as CommentWithImage.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a comment to cell C3 (row index 2, column index 2)
            int commentIdx = worksheet.Comments.Add(2, 2);
            Comment comment = worksheet.Comments[commentIdx];
            comment.Note = "Comment with image background";

            // Load the image file if it exists
            string imagePath = "background.png";
            if (File.Exists(imagePath))
            {
                byte[] imageData = File.ReadAllBytes(imagePath);
                // Set the comment shape's fill to use the loaded image as a texture
                comment.CommentShape.Fill.Type = FillType.Texture;
                comment.CommentShape.Fill.ImageData = imageData;
            }
            else
            {
                Console.WriteLine($"Image file '{imagePath}' not found. The comment will be saved without background image.");
            }

            // Save the workbook
            string outputPath = "CommentWithImage.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
