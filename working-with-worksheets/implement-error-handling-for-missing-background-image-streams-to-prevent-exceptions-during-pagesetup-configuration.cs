// Title: Handle Missing Header Background Image Safely with Aspose.Cells PageSetup (C#)
// Description: Demonstrates how to create a workbook, verify the existence of a PNG file, read its bytes safely, and call PageSetup.SetPicture only when valid data is available. If the image is absent or unreadable, a text header is applied. All steps—including file I/O, picture insertion, and workbook saving—are wrapped in try/catch blocks to prevent unhandled exceptions.
// Keywords: Aspose.Cells | C# | PageSetup.SetPicture | header background image | missing image file | error handling | fallback header text | exception safe workbook | byte array image load | worksheet header picture
// Common Searches: Aspose.Cells set header image missing file | PageSetup SetPicture error handling C# | fallback to text header Aspose.Cells | avoid exception when background image not found | how to check image existence before SetPicture
// Developer Intent: Add robust error handling to prevent exceptions when the background image stream is missing or cannot be read while configuring a worksheet header picture with Aspose.Cells.
// Use Cases: Validate the image path before attempting to load it. | Read the image into a byte array inside a try/catch block and return null on failure. | Invoke PageSetup.SetPicture only when a non‑empty byte array is present. | Provide a default text header as a graceful fallback when the image is unavailable. | Log or display informative messages for missing files, read errors, picture‑setting failures, and save issues.
// AI Prompts: Create a reusable C# method that loads an image file into a byte[] with full exception handling and returns null on failure for use with Aspose.Cells PageSetup.SetPicture. | Show how to replace the hard‑coded fallback header text with a configurable template that can include workbook metadata. | Generate code that writes detailed error information to a log file instead of the console while handling missing background images in Aspose.Cells page setup.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, verify the existence of a PNG file, read its bytes safely, and call PageSetup.SetPicture only when valid data is available. If the image is absent or unreadable, a text header is applied. All steps—including file I/O, picture insertion, and workbook saving—are wrapped in try/catch blocks to prevent unhandled exceptions.
    public class PageSetupBackgroundImageErrorHandlingDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            PageSetup pageSetup = worksheet.PageSetup;

            // Path to the background image (could be missing)
            string imagePath = "header_background.png";

            byte[] imageData = null;

            // Attempt to read the image file safely
            if (File.Exists(imagePath))
            {
                try
                {
                    imageData = File.ReadAllBytes(imagePath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error reading image file: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Image file not found: {imagePath}");
            }

            // Only set the picture if we have valid image data
            if (imageData != null && imageData.Length > 0)
            {
                try
                {
                    // Set picture in the center section of the header (isFirst=false, isEven=false, isHeader=true, section=1)
                    Picture picture = pageSetup.SetPicture(false, false, true, 1, imageData);
                    // Use the picture placeholder in the header
                    pageSetup.SetHeader(1, "&G"); // &G prints the picture
                    Console.WriteLine("Header background image set successfully.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error setting header picture: {ex.Message}");
                }
            }
            else
            {
                // Fallback: set simple text header when image is unavailable
                pageSetup.SetHeader(1, "Sample Header Text");
                Console.WriteLine("Header picture not set; using text header instead.");
            }

            // Save the workbook
            string outputPath = "PageSetupBackgroundImageDemo.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
    }
}
