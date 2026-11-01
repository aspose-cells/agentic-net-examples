// Title: How to extract a worksheet’s background image and save it as a PNG file using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells extract worksheet background image to file | save Excel worksheet background picture as PNG with Aspose.Cells | retrieve background image stream from a worksheet using Aspose.Cells .NET | how to get background image bytes from an Aspose.Cells worksheet | export worksheet background image from .xlsx using Aspose.Cells API
// Tags: worksheet background image extraction Aspose.Cells | save worksheet background as PNG C# | Aspose.Cells retrieve background image bytes | export Excel sheet background image .NET | extract background image stream from worksheet

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, checks a worksheet for a background image, and if present writes the image byte array to a PNG file on disk using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputImagePath = "WorksheetBackground.png";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (or any specific worksheet by index/name)
            Worksheet worksheet = workbook.Worksheets[0];

            // Check if the worksheet contains a background image
            if (worksheet.BackgroundImage != null && worksheet.BackgroundImage.Length > 0)
            {
                // Save the background image bytes to a PNG file
                File.WriteAllBytes(outputImagePath, worksheet.BackgroundImage);
                Console.WriteLine($"Background image saved to: {outputImagePath}");
            }
            else
            {
                Console.WriteLine("The worksheet does not contain a background image.");
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
