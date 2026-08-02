// Title: Add a Footer Watermark Image to Every Worksheet with Aspose.Cells (C#)
// Description: This example demonstrates how to load a PNG file, verify its presence, and use Aspose.Cells' `SetFooterPicture` method to embed the image as a center‑section footer on all worksheets. It also shows how to clear the footer when the image is unavailable and save the workbook.
// Keywords: Aspose.Cells footer picture | C# footer watermark | SetFooterPicture example | apply image to worksheet footer | Aspose.Cells page setup | add logo to Excel footer
// Common Searches: Aspose.Cells add image to footer C# | set footer picture for all sheets Aspose.Cells | C# code to insert watermark in Excel footer | Aspose.Cells page setup footer image multiple worksheets
// Developer Intent: Insert the same picture as a footer watermark across all worksheets in a workbook.
// Use Cases: Brand every sheet with a company logo in the footer. | Include a confidentiality watermark on printable reports. | Attach a copyright image to each worksheet before distribution.
// AI Prompts: Show how to place the footer watermark in the left section instead of the center using Aspose.Cells. | Provide code that reads the image from a stream and applies it as a footer picture to each worksheet. | Explain how to use different footer images for odd and even pages with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This example demonstrates how to load a PNG file, verify its presence, and use Aspose.Cells' `SetFooterPicture` method to embed the image as a center‑section footer on all worksheets. It also shows how to clear the footer when the image is unavailable and save the workbook.
class FooterWatermarkDemo
{
    static void Main()
    {
        try
        {
            // Path to the image file that will be used as the footer watermark
            string imagePath = "watermark.png";

            // Verify that the image file exists before attempting to read it
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                Console.WriteLine("The workbook will be created without a footer watermark.");
            }

            // Load the image file into a byte array (if it exists)
            byte[] imageData = null;
            if (File.Exists(imagePath))
            {
                imageData = File.ReadAllBytes(imagePath);
            }

            // Create a new workbook (or load an existing one if needed)
            Workbook workbook = new Workbook();

            // Apply the footer picture to every worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (imageData != null)
                {
                    // Insert the picture into the center section of the footer (section index 1)
                    Picture picture = sheet.PageSetup.SetFooterPicture(1, imageData);
                    // Set the footer script to display the picture
                    sheet.PageSetup.SetFooter(1, "&G");
                }
                else
                {
                    // Optionally clear any existing footer content
                    sheet.PageSetup.SetFooter(1, string.Empty);
                }
            }

            // Save the workbook with the footer watermarks applied
            string outputPath = "WorkbookWithFooterWatermark.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred:");
            Console.WriteLine(ex.Message);
        }
    }
}
