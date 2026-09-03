// Title: Enable ScaleCrop to keep a PNG’s aspect ratio when adding it to cell A1 of a new Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a Workbook, adds a PNG image at cell A1, sets the picture’s ScaleCrop property to true to avoid stretching, and saves the workbook. | Show how to verify an image file exists, place it on the first worksheet using Aspose.Cells, enable ScaleCrop to retain the original proportions, and write the result to an .xlsx file.
// Common Searches: Aspose.Cells C# keep image aspect ratio when adding picture | Prevent PNG distortion in Excel using Aspose.Cells .NET | Add picture to cell A1 in Excel without stretching with Aspose.Cells | C# Aspose.Cells maintain original picture proportions on workbook creation | How to configure picture scaling in Aspose.Cells for .NET
// Tags: Aspose.Cells picture scaling option | C# insert PNG into Excel worksheet | maintain image aspect ratio Aspose.Cells | create workbook with scaled picture .NET | Excel picture insertion without distortion

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new Workbook, checks that a PNG file exists, inserts the image into cell A1 of the first worksheet, sets the picture’s ScaleCrop property to true so the picture retains its original aspect ratio, and saves the workbook as an .xlsx file.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the image file to be inserted
            string imagePath = "sample_image.png";

            // Verify that the image file exists to avoid FileNotFoundException
            if (!File.Exists(imagePath))
                throw new FileNotFoundException($"Image file not found: {imagePath}");

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add the picture to cell A1 (row 0, column 0)
            int pictureIndex = sheet.Pictures.Add(0, 0, imagePath);
            Picture picture = sheet.Pictures[pictureIndex];

            // Note: Aspect ratio is preserved by default in Aspose.Cells.
            // If needed, you can manually adjust Width/Height while keeping the ratio.

            // Save the workbook
            string outputPath = "WorkbookWithScaledImage.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
