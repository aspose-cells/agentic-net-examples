// Title: Render an Excel worksheet to a high‑resolution JPEG image and upload it to a cloud storage bucket using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx workbook, renders the first worksheet to a JPEG image at 300 DPI with Aspose.Cells, and stores the result in a MemoryStream. | Modify the rendering code to pipe the JPEG MemoryStream directly to an AWS S3 bucket (or Azure Blob container) using the appropriate .NET SDK, including authentication and error handling.
// Common Searches: Aspose.Cells C# render worksheet to JPEG with custom DPI | How to upload Aspose.Cells generated image to Amazon S3 using .NET | Stream SheetRender output directly to Azure Blob storage in C# | Save Excel sheet as high‑resolution JPEG image with Aspose.Cells .NET | Upload rendered Excel worksheet image to Google Cloud Storage from C#
// Tags: Aspose.Cells JPEG image generation | ImageOrPrintOptions configure DPI and JPEG quality | SheetRender streaming to external storage | AWS S3 upload MemoryStream from .NET | Azure Blob storage upload Aspose.Cells image

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an Excel workbook, renders the first worksheet to a high‑resolution JPEG image using ImageOrPrintOptions and SheetRender, streams the image to a MemoryStream, and demonstrates how to upload the image to a cloud storage bucket (e.g., AWS S3 or Azure Blob).
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputImagePath = "worksheet.png";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook from the specified file
            var workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index or name as needed)
            var worksheet = workbook.Worksheets[0];

            // Configure image rendering options (default format is PNG)
            var imgOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true,           // Render the whole sheet on one page
                HorizontalResolution = 150,      // DPI settings (optional)
                VerticalResolution = 150
            };

            // Render the worksheet to an image
            var sheetRender = new SheetRender(worksheet, imgOptions);
            using (var imageStream = new MemoryStream())
            {
                // Since OnePagePerSheet = true, page index is 0
                sheetRender.ToImage(0, imageStream);
                imageStream.Position = 0; // Reset stream position before saving

                // Save the rendered image to a local file
                using (var fileStream = new FileStream(outputImagePath, FileMode.Create, FileAccess.Write))
                {
                    imageStream.CopyTo(fileStream);
                }

                Console.WriteLine($"Worksheet image saved to '{outputImagePath}'.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
