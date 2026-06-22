using System;
using System.IO;
using System.Drawing;                     // For Color
using Aspose.Cells;
using Aspose.Cells.Rendering;            // For ImageOrPrintOptions
using Aspose.Cells.Drawing;              // For ImageType

class ExportRangeAsImage
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "range_image.png";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: File \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Create the range to export (e.g., A1:J25)
            Aspose.Cells.Range range = worksheet.Cells.CreateRange("A1:J25");

            // Set image conversion options (PNG, 300 DPI)
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png,
                HorizontalResolution = 300,
                VerticalResolution = 300
                // BackgroundColor property is not available in this version of Aspose.Cells
            };

            // Convert the range to an image and obtain the byte array
            byte[] imageData = range.ToImage(options);

            // Save the image bytes to a file
            File.WriteAllBytes(outputPath, imageData);

            Console.WriteLine($"Range image saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}