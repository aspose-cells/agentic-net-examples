// Title: Bring a picture shape to the front of the drawing order in an Excel worksheet with Aspose.Cells for .NET (C#)
// Common Searches: set picture ZOrderPosition to front using Aspose.Cells | move image shape to top layer in Excel with Aspose.Cells .NET | change drawing order of worksheet shapes programmatically Aspose.Cells | overlay picture over other shapes in Excel workbook using Aspose.Cells | bring inserted picture to front in Aspose.Cells example
// Tags: Aspose.Cells picture ZOrderPosition | C# shape layering Aspose.Cells | Excel worksheet picture front order | Aspose.Cells drawing order API | set shape front order .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing; // Provides Picture and Shape classes

// The sample loads an existing workbook, inserts a PNG picture into the first worksheet, assigns its ZOrderPosition to the highest value to move it to the front of the drawing order, and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = @"C:\Temp\Input.xlsx";
            string imagePath = @"C:\Temp\MyImage.png";
            string outputPath = @"C:\Temp\Output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Verify that the image file exists
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a picture to the worksheet at row 2, column 1 (zero‑based indices)
            int pictureIndex = sheet.Pictures.Add(2, 1, imagePath);

            // Retrieve the picture object
            Picture picture = sheet.Pictures[pictureIndex];

            // Bring the picture to the front by setting a high Z‑order position
            picture.ZOrderPosition = sheet.Pictures.Count;

            // Save the workbook with the updated drawing order
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
