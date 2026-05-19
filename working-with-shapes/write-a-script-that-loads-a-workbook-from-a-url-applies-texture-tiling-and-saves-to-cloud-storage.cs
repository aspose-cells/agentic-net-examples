using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsTextureTilingDemo
{
    class Program
    {
        static async Task Main(string[] args)
        {
            try
            {
                // URL of the source Excel file
                const string excelUrl = "https://example.com/sample.xlsx";

                // URL of the texture image to be used for tiling
                const string textureImageUrl = "https://example.com/texture.png";

                // Download Excel file into a memory stream
                using (HttpClient httpClient = new HttpClient())
                using (Stream excelStream = await httpClient.GetStreamAsync(excelUrl))
                {
                    // Load workbook from the stream
                    using (Workbook workbook = new Workbook(excelStream))
                    {
                        // Access the first worksheet
                        Worksheet worksheet = workbook.Worksheets[0];

                        // Add a rectangle shape (row, column, top offset, left offset, height, width)
                        Shape rectangle = worksheet.Shapes.AddRectangle(2, 0, 0, 0, 200, 300);

                        // Set fill type to texture
                        rectangle.Fill.FillType = FillType.Texture;

                        // Download texture image bytes
                        byte[] textureBytes = await httpClient.GetByteArrayAsync(textureImageUrl);

                        // Apply texture image data
                        TextureFill textureFill = rectangle.Fill.TextureFill;
                        textureFill.ImageData = textureBytes;

                        // Enable tiling
                        textureFill.IsTiling = true;

                        // Configure tile picture options (scale, offset, etc.)
                        TilePicOption tileOptions = new TilePicOption
                        {
                            ScaleX = 50,   // 50% horizontal scaling
                            ScaleY = 50,   // 50% vertical scaling
                            OffsetX = 10,  // 10 pixels horizontal offset
                            OffsetY = 10   // 10 pixels vertical offset
                        };
                        textureFill.TilePicOption = tileOptions;

                        // Save the modified workbook to a local file
                        const string outputPath = "tiledWorkbook.xlsx";

                        // Ensure we don't overwrite an existing file unintentionally
                        if (File.Exists(outputPath))
                        {
                            File.Delete(outputPath);
                        }

                        workbook.Save(outputPath, SaveFormat.Xlsx);
                        Console.WriteLine($"Workbook saved successfully to: {Path.GetFullPath(outputPath)}");
                    }
                }
            }
            catch (HttpRequestException ex)
            {
                Console.Error.WriteLine($"Network error: {ex.Message}");
            }
            catch (IOException ex)
            {
                Console.Error.WriteLine($"File I/O error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}