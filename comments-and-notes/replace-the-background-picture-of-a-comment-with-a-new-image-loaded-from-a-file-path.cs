// Title: Replace an Excel comment background with an image file using Aspose.Cells for .NET
// Description: Shows how to create a workbook, add a comment to cell B2, read an image file into a byte array, assign it to the comment's shape fill via FillFormat.ImageData, and save the workbook. Includes a file‑existence check and basic exception handling.
// Keywords: Aspose.Cells comment background image | C# set comment picture | FillFormat.ImageData | Excel comment custom background | load image from file Aspose.Cells | replace comment shape fill | Aspose.Cells .NET example | comment background PNG | Excel comment image programmatically | Aspose.Cells API Fill.ImageData
// Common Searches: how to set comment background image Aspose.Cells C# | replace Excel comment picture from file path | Aspose.Cells FillFormat.ImageData example | add custom image to comment shape .NET | change comment background programmatically | load PNG into comment background Aspose.Cells
// Developer Intent: Replace the existing background of a worksheet comment with a new image loaded from a local file.
// Use Cases: Brand a report by placing a company logo as the comment background for key cells. | Provide visual context, such as a small diagram, directly behind a comment to aid data interpretation. | Create a themed template where each comment shares a consistent background image for visual uniformity. | Generate localized worksheets that display region‑specific icons or flags in comment backgrounds.
// AI Prompts: Write C# code that uses Aspose.Cells to set a comment's background image from a file path, including checks for missing or unsupported files. | Explain how to automatically resize a comment shape after assigning a background image so the picture fits without distortion. | Show how to apply the same background image to multiple comments across several worksheets in one workbook using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to create a workbook, add a comment to cell B2, read an image file into a byte array, assign it to the comment's shape fill via FillFormat.ImageData, and save the workbook. Includes a file‑existence check and basic exception handling.
class ReplaceCommentBackground
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a comment to cell B2 (row 1, column 1)
            int commentIdx = worksheet.Comments.Add("B2");
            Comment comment = worksheet.Comments[commentIdx];
            comment.Note = "This comment will have a custom background image.";

            // Path to the new background image
            string imagePath = "newBackground.png";

            // Verify that the image file exists before attempting to read it
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Load the background image into a byte array
            byte[] imageData = File.ReadAllBytes(imagePath);

            // Replace the comment's background picture using FillFormat.ImageData
            comment.CommentShape.Fill.ImageData = imageData;

            // Save the workbook to a file
            workbook.Save("CommentWithNewBackground.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
