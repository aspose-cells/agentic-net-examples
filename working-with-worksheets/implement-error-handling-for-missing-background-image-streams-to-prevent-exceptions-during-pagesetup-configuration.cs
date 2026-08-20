// Title: Safely Set Header Image in Aspose.Cells .NET – Handle Missing Files and Avoid Exceptions
// Description: Demonstrates how to create a workbook, verify the existence of a header image, load its bytes safely, and apply PageSetup.SetPicture only when valid data is available. The example clears the header placeholder when the image is absent and includes comprehensive try/catch blocks to ensure the workbook saves without runtime errors.
// Keywords: Aspose.Cells header image | SetPicture safe usage | missing image handling .NET | page setup header picture error handling | prevent exception Aspose.Cells | C# workbook header graphic | conditional header image Aspose
// Common Searches: Aspose.Cells set header picture only if file exists | avoid exception when image not found in PageSetup.SetPicture | C# check for image before adding to worksheet header | how to skip header graphic in Aspose.Cells if missing | error handling for background image in Aspose.Cells
// Developer Intent: Add robust logic that sets a worksheet header picture only when a valid image stream is present, otherwise clears the placeholder to prevent runtime exceptions.
// Use Cases: Generating reports with an optional company logo that may not be deployed on every server. | Creating workbook templates that conditionally include a header graphic without crashing if the file is absent. | Running batch workbook creation where missing images are logged and processing continues uninterrupted.
// AI Prompts: Write C# code using Aspose.Cells to add a footer picture after confirming the image file exists and handling any I/O errors. | Refactor the sample to extract image loading into a reusable method while guaranteeing the workbook saves even when the image is missing. | Create unit tests for the safe header image handling logic in an Aspose.Cells workbook, covering scenarios with existing and missing image files.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsBackgroundImageHandling
{
    // Demonstrates how to create a workbook, verify the existence of a header image, load its bytes safely, and apply PageSetup.SetPicture only when valid data is available. The example clears the header placeholder when the image is absent and includes comprehensive try/catch blocks to ensure the workbook saves without runtime errors.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add some sample data
            worksheet.Cells["A1"].PutValue("Demo of safe background image handling");

            // Path to the header image (can be missing)
            string imagePath = "header.png";

            // Load image data safely
            byte[] imageData = null;
            if (File.Exists(imagePath))
            {
                try
                {
                    imageData = File.ReadAllBytes(imagePath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to read image file: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Image file \"{imagePath}\" not found. Header picture will be skipped.");
            }

            // Configure page setup
            PageSetup pageSetup = worksheet.PageSetup;

            // If image data is valid, set it as a header picture; otherwise, avoid calling SetPicture
            if (imageData != null && imageData.Length > 0)
            {
                try
                {
                    // Parameters: isFirstPage, isEvenPage, isHeader, section (1 = center), image bytes
                    pageSetup.SetPicture(false, false, true, 1, imageData);
                    // Insert picture placeholder into the header
                    pageSetup.SetHeader(1, "&G");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error while setting header picture: {ex.Message}");
                }
            }
            else
            {
                // Ensure header does not contain a picture placeholder that could cause an exception
                pageSetup.SetHeader(1, string.Empty);
            }

            // Save the workbook
            string outputPath = "WorkbookWithHeader.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
