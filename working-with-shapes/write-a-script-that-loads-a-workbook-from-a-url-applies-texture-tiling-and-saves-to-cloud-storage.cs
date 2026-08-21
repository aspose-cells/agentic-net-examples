// Title: Download an Excel workbook from a URL, apply a tiled texture fill to a shape, and save it with Aspose.Cells (C#)
// Description: C# example that uses HttpClient to fetch an XLSX file, loads it into Aspose.Cells via a MemoryStream, adds a rectangle shape, sets its FillType to a tiled BlueTissuePaper texture (with TilePicOption scaling and offset), and writes the modified workbook to a file or cloud storage.
// Keywords: Aspose.Cells download workbook from URL | C# texture fill shape | tiled texture Aspose.Cells | FillType.Texture C# | TilePicOption scaling offset | HttpClient Excel stream | save workbook to cloud storage | Azure Blob Aspose.Cells | async Excel processing .NET | shape fill pattern Excel
// Common Searches: how to load Excel from a web URL using Aspose.Cells | apply tiled texture fill to a shape in Aspose.Cells C# | save modified workbook to Azure Blob with Aspose.Cells | configure TilePicOption for texture scaling in Aspose.Cells | download and edit Excel file in memory stream C#
// Developer Intent: The developer needs to retrieve an Excel file over HTTP, add a rectangle shape with a repeated texture pattern, and persist the updated workbook for further use or cloud upload.
// Use Cases: Automated branding: download a template, overlay a tiled texture on header shapes, then store the file in cloud storage for distribution. | Marketing asset generation: apply a custom patterned fill to shapes after pulling a base workbook from a web service, then deliver the file to a content‑management system. | Batch styling pipeline: fetch multiple Excel reports, enrich them with textured shapes for visual consistency, and save the results to Azure Blob or Amazon S3. | SaaS onboarding: retrieve a starter workbook, programmatically add visual cues using tiled textures, and provide the customized file to end‑users.
// AI Prompts: Write C# code that downloads an XLSX file from a URL, adds a rectangle with a tiled custom image texture using Aspose.Cells, and uploads the result to Azure Blob Storage. | Refactor the script to use "await using" for all disposable objects, add cancellation support, and replace the built‑in BlueTissuePaper texture with a user‑provided PNG while keeping tiling enabled. | Explain how TilePicOption properties (ScaleX, ScaleY, OffsetX, OffsetY) affect the appearance of a tiled texture fill on a shape in Aspose.Cells, with code snippets for different visual outcomes.

using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# example that uses HttpClient to fetch an XLSX file, loads it into Aspose.Cells via a MemoryStream, adds a rectangle shape, sets its FillType to a tiled BlueTissuePaper texture (with TilePicOption scaling and offset), and writes the modified workbook to a file or cloud storage.
class Program
{
    // Entry point
    static async Task Main()
    {
        try
        {
            // URL of the source Excel file
            string excelUrl = "https://example.com/sample.xlsx";

            // Download the Excel file into a memory stream
            using (HttpClient httpClient = new HttpClient())
            using (Stream downloadStream = await httpClient.GetStreamAsync(excelUrl))
            using (MemoryStream workbookStream = new MemoryStream())
            {
                // Copy downloaded data to a seekable stream
                await downloadStream.CopyToAsync(workbookStream);
                workbookStream.Position = 0; // Reset for reading

                // Load workbook from the stream
                Workbook workbook = new Workbook(workbookStream);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add a rectangle shape to demonstrate texture tiling
                // Parameters: upper left row, upper left column, upper left offset, upper left offset, width, height
                Shape rect = sheet.Shapes.AddRectangle(2, 0, 0, 0, 300, 200);

                // Set fill type to texture
                rect.Fill.FillType = FillType.Texture;

                // Configure texture fill
                TextureFill textureFill = rect.Fill.TextureFill;
                textureFill.Type = TextureType.BlueTissuePaper; // Built‑in texture
                textureFill.IsTiling = true;                    // Enable tiling

                // Optional: configure tile picture options (scale, offset, etc.)
                TilePicOption tileOption = new TilePicOption
                {
                    ScaleX = 50,   // 50% horizontal scaling
                    ScaleY = 50,   // 50% vertical scaling
                    OffsetX = 10,  // 10 pixels horizontal offset
                    OffsetY = 10   // 10 pixels vertical offset
                };
                textureFill.TilePicOption = tileOption;

                // Save the modified workbook to a memory stream in XLSX format
                using (MemoryStream outputStream = new MemoryStream())
                {
                    workbook.Save(outputStream, SaveFormat.Xlsx);
                    outputStream.Position = 0; // Reset for further use

                    // Save to local file (ensure the directory exists)
                    string outputPath = Path.Combine(Environment.CurrentDirectory, "modified.xlsx");
                    File.WriteAllBytes(outputPath, outputStream.ToArray());

                    Console.WriteLine($"Workbook saved successfully to: {outputPath}");
                }
            }
        }
        catch (HttpRequestException ex)
        {
            Console.WriteLine($"Error downloading the Excel file: {ex.Message}");
        }
        catch (IOException ex)
        {
            Console.WriteLine($"File I/O error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
