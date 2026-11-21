// Title: Add a picture to an Excel worksheet and lock its aspect ratio with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add image to worksheet while preserving aspect ratio | how to lock picture dimensions proportionally in Excel using Aspose.Cells | set picture width and auto‑calculate height Aspose.Cells .NET | resize inserted picture without distortion Aspose.Cells C# example | C# Aspose.Cells insert PNG and maintain original proportions
// Tags: insert image into Excel worksheet Aspose.Cells | preserve image proportions during resize C# | calculate picture height from fixed width Aspose.Cells | Excel picture scaling without distortion .NET | add PNG to workbook with proportional sizing

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, verifies a PNG file, inserts it as a picture on the first worksheet, computes the height needed to keep the original aspect ratio based on a fixed width, applies the new dimensions, and saves the file as output.xlsx.
class AddPictureWithAspectLock
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            string imagePath = "sample.png";

            // Verify that the image file exists before adding
            if (File.Exists(imagePath))
            {
                try
                {
                    // Add a picture to the worksheet at row 2, column 1 (zero‑based indexes)
                    int pictureIndex = sheet.Pictures.Add(2, 1, imagePath);
                    Picture picture = sheet.Pictures[pictureIndex];

                    // Preserve aspect ratio manually: calculate new height based on desired width
                    int newWidth = 200; // pixels
                    double aspectRatio = (double)picture.Height / picture.Width;
                    int newHeight = (int)Math.Round(newWidth * aspectRatio);

                    picture.Width = newWidth;
                    picture.Height = newHeight;
                }
                catch (Exception picEx)
                {
                    Console.WriteLine($"Failed to add picture: {picEx.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture insertion.");
            }

            // Save the workbook to a file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
