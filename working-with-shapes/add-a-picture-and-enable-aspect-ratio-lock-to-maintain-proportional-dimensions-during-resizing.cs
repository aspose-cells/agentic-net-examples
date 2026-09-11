// Title: Insert a PNG picture into an Excel worksheet and lock its aspect ratio while resizing using Aspose.Cells for .NET
// AI Prompts: Add a picture to a specific cell in a worksheet, set a target width, and automatically compute the height to keep the original proportion with Aspose.Cells. | Write a helper method that accepts any image path and a desired width, inserts the image, and locks its dimensions to preserve the aspect ratio. | Modify the example to specify a fixed height and calculate the matching width so the picture scales proportionally.
// Common Searches: Aspose.Cells C# insert PNG and keep original aspect ratio | how to resize an Excel picture proportionally with Aspose.Cells .NET | preserve image dimensions when adding a picture to a worksheet using Aspose.Cells
// Tags: add picture to worksheet Aspose.Cells | lock picture dimensions Aspose.Cells | compute proportional height Aspose.Cells | resize picture proportionally Aspose.Cells | insert PNG file into Excel workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads a PNG file, inserts it at cell B5 of the first worksheet, resizes the picture to a width of 200 points, calculates the matching height to maintain the original aspect ratio, and saves the workbook as ResultWithPicture.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the image file to be inserted
            string imagePath = @"C:\Images\sample.png";

            // Verify that the image file exists to avoid FileNotFoundException
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Desired width in points (1 point = 1/72 inch)
            const int desiredWidth = 200;

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add the picture to the worksheet at row 5, column 2 (zero‑based indexes)
            int pictureIndex;
            try
            {
                pictureIndex = sheet.Pictures.Add(4, 1, imagePath);
            }
            catch (Exception picEx)
            {
                Console.WriteLine($"Failed to add picture: {picEx.Message}");
                return;
            }

            // Retrieve the added picture object
            Picture picture = sheet.Pictures[pictureIndex];

            // Preserve aspect ratio using the picture's original dimensions
            double originalWidth = picture.Width;
            double originalHeight = picture.Height;

            // Calculate height to maintain aspect ratio
            int calculatedHeight = (int)Math.Round(desiredWidth * originalHeight / originalWidth);

            // Apply the new size
            picture.Width = desiredWidth;
            picture.Height = calculatedHeight;

            // Save the workbook
            string outputPath = "ResultWithPicture.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
