// Title: How to extract and save a worksheet background image as a PNG using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, reads the BackgroundImage byte array from a chosen worksheet, and writes it to a PNG file on disk. | Adapt the sample to export the worksheet background image to a JPEG file and accept the output path as a method parameter. | Create a reusable C# method that returns the background image stream of any worksheet and saves it to a specified image format (PNG, JPEG) using Aspose.Cells.
// Common Searches: Aspose.Cells C# extract worksheet background image to file | save Excel sheet background picture as PNG using Aspose.Cells | retrieve BackgroundImage bytes from a worksheet with Aspose.Cells .NET | how to export worksheet background image from .xlsx using Aspose.Cells
// Tags: Aspose.Cells worksheet background image extraction | C# save worksheet background as PNG | Aspose.Cells export background picture to file | retrieve BackgroundImage byte array Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook, accesses a worksheet, checks for a background image via the BackgroundImage property, and if present writes the raw image bytes to a PNG file, handling missing files and I/O errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (change index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Check if the worksheet contains a background image
            if (sheet.BackgroundImage != null && sheet.BackgroundImage.Length > 0)
            {
                const string outputPath = "WorksheetBackground.png";

                try
                {
                    // Save the raw image bytes directly to a file
                    File.WriteAllBytes(outputPath, sheet.BackgroundImage);
                    Console.WriteLine($"Background image saved to '{outputPath}'.");
                }
                catch (Exception imgEx)
                {
                    Console.WriteLine($"Failed to save background image: {imgEx.Message}");
                }
            }
            else
            {
                Console.WriteLine("The worksheet does not contain a background image.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
