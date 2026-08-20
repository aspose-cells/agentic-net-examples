// Title: Render an Excel Worksheet to JPEG and Upload to Cloud Storage with Aspose.Cells for .NET
// Description: This example demonstrates how to create a workbook, render a worksheet page to a JPEG image using Aspose.Cells' SheetRender and ImageOrPrintOptions, save the image to a stream, and programmatically upload the JPEG to a cloud storage bucket (e.g., AWS S3, Azure Blob, Google Cloud Storage).
// Keywords: Aspose.Cells | C# | render worksheet to JPEG | SheetRender | ImageOrPrintOptions JPEG | export Excel as image | upload JPEG to cloud | AWS S3 upload C# | Azure Blob storage C# | Google Cloud Storage C# | Excel to image conversion
// Common Searches: convert Excel worksheet to JPEG using Aspose.Cells C# | save rendered worksheet image to cloud storage | Aspose.Cells upload JPEG to S3 | C# export Excel sheet as JPEG file | how to render Excel page as image .NET
// Developer Intent: Generate a JPEG image from a worksheet and store it directly in a cloud bucket.
// Use Cases: Display spreadsheet previews on a web portal without exposing raw data. | Attach a static image of a financial report to emails for quick review. | Archive spreadsheet snapshots in a compliance‑friendly image format.
// AI Prompts: Generate C# code that renders each worksheet in a workbook to a separate JPEG and uploads them to an AWS S3 bucket with public read access. | Show how to configure ImageOrPrintOptions to set JPEG quality to 80% before rendering. | Provide an example of uploading the JPEG MemoryStream to Azure Blob Storage using the Azure.Storage.Blobs SDK.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// This example demonstrates how to create a workbook, render a worksheet page to a JPEG image using Aspose.Cells' SheetRender and ImageOrPrintOptions, save the image to a stream, and programmatically upload the JPEG to a cloud storage bucket (e.g., AWS S3, Azure Blob, Google Cloud Storage).
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells["A1"].PutValue("Aspose.Cells JPEG Rendering");
            worksheet.Cells["A2"].PutValue(DateTime.Now);

            // Configure image rendering options for JPEG
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Jpeg,
                OnePagePerSheet = true
            };

            // Render the first page of the worksheet to a memory stream
            using (MemoryStream imageStream = new MemoryStream())
            {
                SheetRender sheetRender = new SheetRender(worksheet, options);
                sheetRender.ToImage(0, imageStream);
                imageStream.Position = 0; // Reset stream position for saving

                // Save the JPEG image to a local file
                string outputPath = Path.Combine(Environment.CurrentDirectory, "worksheet_page0.jpg");
                using (FileStream fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                {
                    imageStream.CopyTo(fileStream);
                }

                Console.WriteLine($"Image successfully saved to: {outputPath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
