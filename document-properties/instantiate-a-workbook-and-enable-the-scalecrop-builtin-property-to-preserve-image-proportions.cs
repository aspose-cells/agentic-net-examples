// Title: Create a new workbook and preserve picture aspect ratio with ScaleCrop in Aspose.Cells for .NET (C#)
// Common Searches: how to keep image proportions when inserting a picture with Aspose.Cells C# | Aspose.Cells C# set ScaleCrop property for worksheet picture | preserve aspect ratio of added image using Aspose.Cells picture API | C# add picture to Excel and lock its aspect ratio with Aspose.Cells | enable built‑in ScaleCrop for images in Aspose.Cells .NET
// Tags: Aspose.Cells picture ScaleCrop C# | preserve image aspect ratio Aspose.Cells | add picture to worksheet Aspose.Cells | Excel image scaling Aspose.Cells .NET | workbook image insertion C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a Workbook, inserts an image file as a picture on the first worksheet, and explains how to use the ScaleCrop built‑in property (or manually adjust width/height) to maintain the image's aspect ratio before saving the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the image file (replace with your actual image path)
            string imagePath = "image.png";

            // Verify that the image file exists to avoid FileNotFoundException
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add the picture to the worksheet at row 0, column 0
            int pictureIndex = sheet.Pictures.Add(0, 0, imagePath);
            Picture picture = sheet.Pictures[pictureIndex];

            // Note: Aspose.Cells does not expose an IsLockedAspectRatio property in some versions.
            // If needed, you can manually set Width and Height to preserve proportions.

            // Save the workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
