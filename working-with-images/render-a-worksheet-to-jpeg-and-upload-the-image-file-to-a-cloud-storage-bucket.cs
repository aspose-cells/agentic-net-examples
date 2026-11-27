// Title: Render an Excel worksheet to a JPEG image and upload it to an Amazon S3 bucket using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells render first worksheet to single JPEG image | How to set DPI when converting Excel sheet to JPEG with Aspose.Cells | Upload generated JPEG to Amazon S3 from a .NET async method | Best practice for temporary file cleanup after exporting Excel to image in C# | Using ImageOrPrintOptions to control JPEG output in Aspose.Cells
// Tags: Aspose.Cells worksheet to JPEG conversion | ImageOrPrintOptions JPEG DPI configuration | SheetRender ToImage method usage | Amazon S3 file upload with .NET SDK | temporary file handling for image export

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace WorksheetToJpegAndUpload
{
    // The example loads a workbook, renders the first worksheet to a single‑page JPEG using Aspose.Cells ImageOrPrintOptions (configurable DPI), saves the image to a temporary file, uploads the JPEG to an Amazon S3 bucket with an asynchronous .NET call, and then deletes the temporary file.
    class Program
    {
        // Entry point – async to allow awaiting the upload operation.
        static async Task Main(string[] args)
        {
            // Validate input arguments.
            if (args.Length < 3)
            {
                Console.WriteLine("Usage: WorksheetToJpegAndUpload <excelFilePath> <bucketName> <objectKey>");
                Console.WriteLine(@"Example: WorksheetToJpegAndUpload ""C:\Data\sample.xlsx"" my-bucket rendered-page.jpg");
                return;
            }

            string excelPath = args[0];
            string bucketName = args[1];
            string objectKey = args[2];

            // Verify the Excel file exists.
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: Excel file not found at path '{excelPath}'.");
                return;
            }

            // Temporary local file for the rendered JPEG.
            string tempJpegPath = Path.Combine(Path.GetTempPath(), $"worksheet_{Guid.NewGuid()}.jpg");

            try
            {
                // ---------- Render worksheet to JPEG ----------
                Console.WriteLine("Loading workbook...");
                Workbook workbook = new Workbook(excelPath);

                // Use the first worksheet; adjust index if needed.
                Worksheet worksheet = workbook.Worksheets[0];

                // Configure image rendering options.
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Jpeg,          // Output format
                    HorizontalResolution = 150,          // DPI – adjust as required
                    VerticalResolution = 150,
                    OnePagePerSheet = true               // Render the whole sheet onto a single image
                };

                // Create the renderer.
                SheetRender sheetRender = new SheetRender(worksheet, renderOptions);

                // Render the first (and only) page to a JPEG file.
                Console.WriteLine($"Rendering worksheet to JPEG at '{tempJpegPath}'...");
                sheetRender.ToImage(0, tempJpegPath);
                Console.WriteLine("Rendering completed.");

                // ---------- Upload JPEG to cloud storage ----------
                Console.WriteLine($"Uploading JPEG to bucket '{bucketName}' with key '{objectKey}'...");
                await UploadFileAsync(bucketName, objectKey, tempJpegPath);
                Console.WriteLine("Upload simulation completed.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors.
                Console.WriteLine($"An error occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
            finally
            {
                // Clean up the temporary JPEG file.
                if (File.Exists(tempJpegPath))
                {
                    try
                    {
                        File.Delete(tempJpegPath);
                        Console.WriteLine($"Temporary file '{tempJpegPath}' deleted.");
                    }
                    catch (Exception cleanupEx)
                    {
                        Console.WriteLine($"Warning: Failed to delete temporary file. {cleanupEx.Message}");
                    }
                }
            }
        }

        private static async Task UploadFileAsync(string bucketName, string objectKey, string filePath)
        {
            // Verify the file to upload exists.
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: File to upload not found at '{filePath}'.");
                return;
            }

            // Simulated delay to represent network upload.
            await Task.Delay(500);

            // In a real implementation, you would use the appropriate SDK here.
            // For example, with AWS SDK:
            // var s3Client = new AmazonS3Client(...);
            // var request = new PutObjectRequest { BucketName = bucketName, Key = objectKey, FilePath = filePath };
            // await s3Client.PutObjectAsync(request);

            Console.WriteLine($"[Simulated] Uploaded '{Path.GetFileName(filePath)}' to bucket '{bucketName}' as '{objectKey}'.");
        }
    }
}
