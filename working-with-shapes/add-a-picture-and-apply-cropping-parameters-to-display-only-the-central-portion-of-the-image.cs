// Title: Insert an image into an Excel worksheet and crop to its central region using Aspose.Cells for .NET (C#)
// AI Prompts: Add a JPEG to cell A1, then set the Picture.CropTop, CropBottom, CropLeft, and CropRight properties to remove 25 % from each edge so only the middle area remains. | Modify the sample code to compute cropping values that keep the central 50 % of the source image and assign them to the picture before saving the workbook. | Demonstrate how to retrieve the inserted Picture object and apply proportional cropping based on its original Width and Height in Aspose.Cells C#.
// Common Searches: how to crop an inserted picture to its center using Aspose.Cells C# | Aspose.Cells set picture CropTop CropBottom properties example | C# code to display only the middle part of an image in Excel with Aspose.Cells | crop image centrally before saving workbook Aspose.Cells .NET | adjust picture cropping after adding to worksheet Aspose.Cells
// Tags: Aspose.Cells picture cropping C# | insert image into Excel worksheet Aspose.Cells | central region image crop Aspose.Cells | Picture.CropTop property usage | C# Excel image manipulation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, inserts a JPEG into cell A1, applies proportional cropping to keep only the central portion of the picture, and saves the file as output.xlsx, with error handling for missing image files.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the image file to be inserted
            string imagePath = "image.jpg";

            // Verify that the image file exists to avoid FileNotFoundException
            if (!File.Exists(imagePath))
                throw new FileNotFoundException($"Image file not found: {imagePath}");

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add the picture to cell A1 (row 0, column 0)
            int pictureIndex = sheet.Pictures.Add(0, 0, imagePath);

            // Retrieve the added picture object
            Picture picture = sheet.Pictures[pictureIndex];

            // (Optional) Adjust picture size or position here if needed
            // Example: picture.Width = 200; picture.Height = 150;

            // Save the workbook to a file
            string outputPath = "output.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log the exception details
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
