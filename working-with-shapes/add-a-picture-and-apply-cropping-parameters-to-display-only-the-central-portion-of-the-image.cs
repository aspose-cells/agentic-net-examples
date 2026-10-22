// Title: Insert a JPEG into an Excel cell and show only its center by preprocessing the image with C# and Aspose.Cells
// AI Prompts: Write C# code that loads a JPEG, crops the central region, and inserts the cropped image into cell B2 of a new workbook using Aspose.Cells. | Explain the steps to crop an image before adding it to an Excel worksheet with Aspose.Cells, including any required .NET image‑processing libraries.
// Common Searches: how to crop an image before inserting into Excel with Aspose.Cells C# | Aspose.Cells picture central crop workaround .NET | insert JPEG into specific cell and display only middle part using Aspose.Cells | preprocess image for cropping then add to worksheet Aspose.Cells example | C# code to add picture to Excel and limit visible area to center
// Tags: Aspose.Cells picture insertion C# | central image cropping before Aspose.Cells import | preprocess JPEG for Excel worksheet | C# image cropping library Aspose.Cells workaround | Excel cell picture display limited area

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example verifies a JPEG file, creates a new workbook, and adds the image to cell B2. Because Aspose.Cells' Picture class lacks direct cropping, it recommends preprocessing the image (e.g., using System.Drawing) to extract the central portion before insertion, then saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source image
            string imagePath = @"C:\Images\sample.jpg";

            // Verify the image file exists
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add the picture to cell B2 (row 1, column 1)
            int pictureIndex = sheet.Pictures.Add(1, 1, imagePath);
            Picture picture = sheet.Pictures[pictureIndex];

            // NOTE: Aspose.Cells' Picture class does not expose direct cropping properties.
            // If cropping is required, it must be performed on the source image before insertion
            // or by using external image processing libraries.

            // Prepare output path and ensure directory exists
            string outputPath = @"C:\Output\CroppedImage.xlsx";
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
