// Title: Set Worksheet Background Image in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to verify an image file, load it into a byte array, assign it to the Worksheet.BackgroundImage property, and save the workbook. Includes basic error handling for missing files.
// Keywords: Aspose.Cells worksheet background image | C# set worksheet background picture | Excel background picture Aspose.Cells | Worksheet.BackgroundImage property | load image bytes C# Aspose | apply branding to Excel sheet | programmatic Excel background image | Aspose.Cells .NET background image example
// Common Searches: how to set worksheet background image using Aspose.Cells C# | Aspose.Cells add background picture to Excel sheet .NET | C# code example for worksheet background image Aspose | set Excel worksheet background from file programmatically | Aspose.Cells background image tutorial
// Developer Intent: Apply a specific image file as the background of an Excel worksheet using Aspose.Cells for .NET.
// Use Cases: Brand corporate reports by overlaying a logo or watermark on each sheet. | Create reusable workbook templates that include a predefined visual background. | Design dashboards with a consistent theme across multiple worksheets. | Generate printable forms that require a letterhead or background graphic. | Automate production of worksheets that need a company‑wide visual identity.
// AI Prompts: Generate C# code to set a PNG file as the background of a selected worksheet with Aspose.Cells. | Show how to replace an existing worksheet background image with a new one in an existing workbook. | Provide robust error handling for missing or unsupported image files when applying a worksheet background. | Explain how to remove a background image from a worksheet using Aspose.Cells. | Demonstrate applying the same background image to all worksheets in a workbook with a loop.

using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetBackgroundDemo
{
    // Demonstrates how to verify an image file, load it into a byte array, assign it to the Worksheet.BackgroundImage property, and save the workbook. Includes basic error handling for missing files.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the background image file
                string imagePath = "background.jpg";

                // Verify that the image file exists before attempting to read it
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file not found: {imagePath}");
                    return;
                }

                // Load image data into a byte array
                byte[] imageData = File.ReadAllBytes(imagePath);

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Apply the background image to the worksheet
                worksheet.BackgroundImage = imageData;

                // Save the workbook
                string outputPath = "WorksheetWithBackground.xlsx";
                workbook.Save(outputPath);

                Console.WriteLine($"Worksheet background image applied successfully. Saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
