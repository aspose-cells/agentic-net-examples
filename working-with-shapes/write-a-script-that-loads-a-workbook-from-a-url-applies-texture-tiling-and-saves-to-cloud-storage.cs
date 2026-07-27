// Title: C# – Download Excel from URL, apply tiled texture fill to a shape, and save to cloud storage with Aspose.Cells
// Description: Download an Excel workbook via HttpClient, load it into an Aspose.Cells Workbook, add a rectangle shape with a tiled BlueTissuePaper texture (custom TilePicOption), then write the file to a MemoryStream and upload it to Azure Blob or other cloud storage.
// Keywords: Aspose.Cells texture fill | C# download Excel from URL | tiled texture shape Aspose.Cells | TilePicOption scale offset | save workbook to Azure Blob | .NET Excel shape fill | cloud storage Aspose.Cells example
// Common Searches: how to apply texture tiling to a shape using Aspose.Cells .NET | download Excel file from web URL and modify with Aspose.Cells | save Aspose.Cells workbook to Azure Blob storage | configure TilePicOption for texture fill in C# | Aspose.Cells example for shape fill and cloud upload
// Developer Intent: Load an Excel file from a remote URL, add a rectangle with a tiled texture fill, and persist the modified workbook to cloud storage using Aspose.Cells for .NET.
// Use Cases: Automated report generation that adds a textured banner to each downloaded template before storing the result in Azure Blob. | Batch processing of client spreadsheets to embed a tiled watermark texture and archive the files in Amazon S3 or Google Cloud Storage. | Dynamic creation of marketing dashboards where a background shape with a custom texture is applied after pulling the base workbook from a web service.
// AI Prompts: Generate C# code that fetches an Excel file from a URL, inserts a rectangle shape with a tiled BlueTissuePaper texture using Aspose.Cells, and uploads the result to Azure Blob storage. | Show how to set TilePicOption properties (ScaleX, ScaleY, OffsetX, OffsetY) for a texture fill on a shape in Aspose.Cells and save the workbook to a MemoryStream.

using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsTextureTilingDemo
{
    // Download an Excel workbook via HttpClient, load it into an Aspose.Cells Workbook, add a rectangle shape with a tiled BlueTissuePaper texture (custom TilePicOption), then write the file to a MemoryStream and upload it to Azure Blob or other cloud storage.
    class Program
    {
        static async Task Main(string[] args)
        {
            try
            {
                // URL of the source Excel file
                string excelUrl = "https://example.com/sample.xlsx";

                // Download the workbook into a stream
                using (HttpClient httpClient = new HttpClient())
                using (Stream downloadStream = await httpClient.GetStreamAsync(excelUrl))
                {
                    // Load the workbook from the downloaded stream
                    using (Workbook workbook = new Workbook(downloadStream))
                    {
                        // Apply texture tiling to a rectangle shape
                        ApplyTextureTiling(workbook);

                        // Save the modified workbook to a memory stream
                        using (MemoryStream outputStream = new MemoryStream())
                        {
                            workbook.Save(outputStream, SaveFormat.Xlsx);
                            outputStream.Position = 0; // Reset for reading

                            // Save the stream to a local file (fallback for environments without Azure SDK)
                            await SaveToFileAsync("modified_workbook.xlsx", outputStream);
                        }
                    }
                }

                Console.WriteLine("Workbook processed and saved successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        private static void ApplyTextureTiling(Workbook workbook)
        {
            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape (position: row 2, column 2, width 200, height 100)
            Shape rect = sheet.Shapes.AddRectangle(2, 2, 0, 0, 200, 100);

            // Set the fill type to texture
            rect.Fill.FillType = FillType.Texture;

            // Use a built‑in texture type
            rect.Fill.TextureFill.Type = TextureType.BlueTissuePaper;

            // Enable tiling
            rect.Fill.TextureFill.IsTiling = true;

            // Configure tile picture options (scale, offset, etc.)
            TilePicOption tileOptions = new TilePicOption
            {
                ScaleX = 50,   // 50% horizontal scaling
                ScaleY = 50,   // 50% vertical scaling
                OffsetX = 10,  // 10 pixels horizontal offset
                OffsetY = 10   // 10 pixels vertical offset
            };
            rect.Fill.TextureFill.TilePicOption = tileOptions;
        }

        private static async Task SaveToFileAsync(string fileName, Stream dataStream)
        {
            try
            {
                // Ensure the target directory exists
                string directory = Path.GetDirectoryName(Path.GetFullPath(fileName));
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Write the stream to a file (overwrite if it exists)
                using (FileStream fileStream = new FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    await dataStream.CopyToAsync(fileStream);
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Failed to save file '{fileName}': {ex.Message}");
                throw;
            }
        }
    }
}
