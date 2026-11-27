// Title: Replace an Excel comment's background image with a PNG file using Aspose.Cells for .NET (C#)
// Common Searches: c# aspocells replace comment background image from png file | how to set a picture as the background of an Excel comment using Aspose.Cells | programmatically change Excel comment background picture in .NET | assign dynamic BackgroundPicture property to a comment in Aspose.Cells | load external image and use it as comment background in an Aspose.Cells workbook
// Tags: comment background picture Aspose.Cells | assign picture to comment background .NET | load image file as comment background C# | dynamic binding for comment BackgroundPicture | replace Excel comment image programmatically | use Pictures.Add to set comment background Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads a workbook, ensures a comment exists on cell A1, adds a PNG picture from a file path, assigns that picture to the comment's BackgroundPicture via a dynamic reference, and saves the workbook with the updated comment background.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string picturePath = "newBackground.png";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input workbook not found: {inputPath}");

            // Verify that the background picture exists
            if (!File.Exists(picturePath))
                throw new FileNotFoundException($"Background picture not found: {picturePath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure a comment exists on cell A1 (create if missing)
            int commentIdx = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIdx];
            comment.Author = "Author";
            comment.Note = string.Empty; // empty comment text

            // Add the picture to the worksheet
            int pictureIdx = sheet.Pictures.Add(0, 0, picturePath);
            Picture picture = sheet.Pictures[pictureIdx];

            // Assign the picture as the comment's background (using dynamic to avoid compile‑time binding issues)
            dynamic dynComment = comment;
            dynComment.BackgroundPicture = picture;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
