// Title: Set a Picture as the Background of an Excel Comment with Aspose.Cells for .NET (C#)
// Description: This example demonstrates how to load an image file into an Aspose.Cells Picture object, assign it to the comment shape via Shape.Fill.Picture, optionally resize the shape, and save the workbook. It includes error handling for missing files and save failures.
// Keywords: Aspose.Cells comment background picture | C# set comment shape fill image | Shape.Fill.Picture Aspose.Cells | embed image in Excel comment .NET | load picture object for comment background
// Common Searches: how to add a picture to an Excel comment using Aspose.Cells | C# Aspose.Cells set comment shape fill to image | Shape.Fill.Picture example Aspose.Cells | embed PNG as comment background in .xlsx | Aspose.Cells comment image background code
// Developer Intent: Apply a custom picture as the background of a cell comment in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Display a company logo behind comment text for branding. | Show thumbnail previews of related charts inside comments. | Create visually rich annotations with custom graphics in generated reports.
// AI Prompts: Generate C# code that loads a PNG file into an Aspose.Cells Picture object and sets it as the background of a comment shape using Shape.Fill.Picture. | Provide a sample that checks for the image file, resizes the comment shape to match the picture dimensions, and saves the workbook as XLSX. | Explain the difference between using Shape.Fill.ImageData and Shape.Fill.Picture for embedding images in comment shapes with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This example demonstrates how to load an image file into an Aspose.Cells Picture object, assign it to the comment shape via Shape.Fill.Picture, optionally resize the shape, and save the workbook. It includes error handling for missing files and save failures.
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

            // Set the comment shape's background image if the file exists
            string imagePath = "background.png"; // replace with your image file path
            if (File.Exists(imagePath))
            {
                try
                {
                    // Load image bytes and assign to the shape's fill image data
                    byte[] imageBytes = File.ReadAllBytes(imagePath);
                    comment.CommentShape.Fill.ImageData = imageBytes;
                }
                catch (Exception imgEx)
                {
                    Console.WriteLine($"Failed to load image '{imagePath}': {imgEx.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Image file not found: {imagePath}. The comment will be saved without a background image.");
            }

            // Optionally adjust the comment shape size
            comment.CommentShape.Width = 250;
            comment.CommentShape.Height = 150;

            // Save the workbook
            string outputPath = "CommentWithImage.xlsx";
            try
            {
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
