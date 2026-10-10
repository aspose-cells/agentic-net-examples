// Title: How to insert an image into an Excel worksheet with Aspose.Cells for .NET, set its height to 200 points, and keep the original aspect ratio
// AI Prompts: Insert a PNG image from a file stream into cell A1 of a new workbook and set only the Height property to 200 points using Aspose.Cells, letting the width adjust automatically. | Add a picture to a worksheet after verifying the source file exists, then save the workbook as output.xlsx while preserving the image's aspect ratio. | Create a workbook, place an image at the top‑left cell, assign picture.Height = 200, and rely on Aspose.Cells to maintain proportional scaling.
// Common Searches: Aspose.Cells C# set picture height points and keep aspect ratio | Insert image into Excel cell without stretching using Aspose.Cells .NET | How to add a PNG to a worksheet and auto‑scale width based on height in Aspose.Cells | C# Aspose.Cells picture dimensions maintain original proportions
// Tags: add picture to worksheet Aspose.Cells | picture height 200 points Aspose.Cells | preserve image aspect ratio Aspose.Cells | insert image from filestream Aspose.Cells .NET | auto scale picture width Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, checks that image.png exists, inserts the image into cell A1, sets picture.Height = 200 points (width scales automatically to preserve the original aspect ratio), and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Verify that the image file exists to avoid FileNotFoundException
            string imagePath = "image.png";
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Insert the picture and configure it
            using (FileStream imgStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
            {
                // Add picture to cell A1 (row 0, column 0)
                int pictureIndex = sheet.Pictures.Add(0, 0, imgStream);
                Picture picture = sheet.Pictures[pictureIndex];

                // Set the height to 200 points; width adjusts proportionally
                picture.Height = 200;
                // Note: Aspect ratio is maintained automatically when only one dimension is set.
            }

            // Save the workbook
            workbook.Save("output.xlsx");
            Console.WriteLine("Workbook saved successfully as output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
