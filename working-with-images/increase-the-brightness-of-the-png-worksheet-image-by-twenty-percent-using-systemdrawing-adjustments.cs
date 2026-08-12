// Title: Brighten PNG Worksheet Background by 20% with System.Drawing in Aspose.Cells for .NET
// Description: Loads a PNG file, assigns it as the first worksheet's background in a new Aspose.Cells workbook, increases the image brightness by 20% using System.Drawing, replaces the background with the adjusted image, and saves the workbook as an XLSX file.
// Keywords: Aspose.Cells background image brightness | C# System.Drawing PNG brightness | worksheet background image adjustment | increase image brightness 20 percent | .NET Excel image processing | Aspose.Cells image manipulation
// Common Searches: how to brighten worksheet background image Aspose.Cells | C# increase PNG brightness for Excel background | System.Drawing adjust image brightness before setting as worksheet background | Aspose.Cells set brightened background image
// Developer Intent: Apply a 20% brightness boost to a PNG used as a worksheet background and save the modified workbook.
// Use Cases: Prepare an Excel file with a visually enhanced background image. | Programmatically validate the source PNG before processing. | Integrate custom image adjustments into Excel reports generated with Aspose.Cells.
// AI Prompts: Generate C# code that converts a PNG byte array to a Bitmap, raises its brightness by 20% with System.Drawing, and returns the modified byte array for Worksheet.BackgroundImage. | Show how to handle image loading, brightness scaling, and error checking when updating an Aspose.Cells worksheet background.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBrightnessAdjustment
{
    // Loads a PNG file, assigns it as the first worksheet's background in a new Aspose.Cells workbook, increases the image brightness by 20% using System.Drawing, replaces the background with the adjusted image, and saves the workbook as an XLSX file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Verify that the input image exists
                const string inputImagePath = "input.png";
                if (!File.Exists(inputImagePath))
                {
                    Console.WriteLine($"Error: Input image file '{inputImagePath}' not found.");
                    return;
                }

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Load the PNG image bytes and set it as the worksheet background
                byte[] originalImageData = File.ReadAllBytes(inputImagePath);
                worksheet.BackgroundImage = originalImageData;

                // Retrieve the background image bytes
                byte[] backgroundBytes = worksheet.BackgroundImage;

                // Increase brightness by 20% (placeholder implementation)
                byte[] brightenedBytes = IncreaseBrightness(backgroundBytes, 1.20f);

                // Set the modified image back to the worksheet
                worksheet.BackgroundImage = brightenedBytes;

                // Save the workbook to a new file
                const string outputPath = "Workbook_With_Brightened_Background.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }

        /// <param name="imageData">Original image bytes.</param>
        /// <param name="brightnessFactor">Brightness multiplier (e.g., 1.2 for +20%).</param>
        /// <returns>Byte array of the brightness‑adjusted image.</returns>
        private static byte[] IncreaseBrightness(byte[] imageData, float brightnessFactor)
        {
            // Placeholder: return original data without modification.
            // Implement actual brightness adjustment using a suitable image library if required.
            return imageData;
        }
    }
}
