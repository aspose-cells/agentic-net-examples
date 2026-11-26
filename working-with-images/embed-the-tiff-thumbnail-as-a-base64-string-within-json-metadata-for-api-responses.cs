// Title: Create a TIFF thumbnail of an Excel worksheet and embed it as a Base64 string in JSON using Aspose.Cells for .NET
// Common Searches: how to generate a TIFF thumbnail from an Excel sheet with Aspose.Cells in C# | embed a base64 encoded TIFF image in JSON response using Aspose.Cells .NET | render the first worksheet to TIFF and include it in an API payload with Aspose.Cells | convert an Excel worksheet thumbnail to a Base64 string for a web API in C# | Aspose.Cells C# create JSON metadata that contains a sheet thumbnail
// Tags: Aspose.Cells render worksheet to TIFF | C# convert TIFF to Base64 | JSON metadata with embedded image | SheetRender TIFF thumbnail .NET | ImageOrPrintOptions TIFF compression LZW | Aspose.Cells generate sheet thumbnail for API

using System;
using System.IO;
using System.Text;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsThumbnailJson
{
    // The example creates a workbook, populates it with sample data, and uses Aspose.Cells SheetRender with ImageOrPrintOptions to render the first worksheet to a TIFF image in memory. The TIFF bytes are then converted to a Base64 string, which is placed into a JSON object alongside metadata such as file name, sheet name, thumbnail format, generation timestamp, size, and the image options used. The JSON is written to the console and saved to a file, and the original workbook is saved as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a sample workbook (or load an existing one)
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "DemoSheet";

                // Populate some sample data
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Quantity");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(85);
                sheet.Cells["A4"].PutValue("Cherry");
                sheet.Cells["B4"].PutValue(60);

                // -----------------------------------------------------------------
                // 2. Render the first sheet to a TIFF thumbnail in memory
                // -----------------------------------------------------------------
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Tiff,
                    HorizontalResolution = 96,
                    VerticalResolution = 96,
                    Transparent = true,
                    TiffCompression = TiffCompression.CompressionLZW,
                    TiffColorDepth = ColorDepth.Format24bpp
                };

                // SheetRender can render a specific sheet index; we use 0 for the first sheet
                SheetRender renderer = new SheetRender(sheet, imgOptions);

                using (MemoryStream tiffStream = new MemoryStream())
                {
                    // Render to the memory stream
                    renderer.ToImage(0, tiffStream);

                    // Ensure the stream is positioned at the beginning before reading
                    tiffStream.Position = 0;

                    // -----------------------------------------------------------------
                    // 3. Convert the TIFF bytes to a Base64 string
                    // -----------------------------------------------------------------
                    byte[] tiffBytes = tiffStream.ToArray();
                    string base64Thumbnail = Convert.ToBase64String(tiffBytes);

                    // -----------------------------------------------------------------
                    // 4. Build JSON metadata containing the Base64 thumbnail
                    // -----------------------------------------------------------------
                    var metadata = new
                    {
                        FileName = "DemoWorkbook.xlsx",
                        SheetName = sheet.Name,
                        ThumbnailFormat = "tiff",
                        ThumbnailBase64 = base64Thumbnail,
                        GeneratedOnUtc = DateTime.UtcNow.ToString("o"),
                        ThumbnailSizeBytes = tiffBytes.Length,
                        ImageOptions = new
                        {
                            imgOptions.ImageType,
                            imgOptions.HorizontalResolution,
                            imgOptions.VerticalResolution,
                            imgOptions.Transparent,
                            imgOptions.TiffCompression,
                            imgOptions.TiffColorDepth
                        }
                    };

                    // Serialize with indentation for readability
                    string jsonOutput = JsonSerializer.Serialize(metadata, new JsonSerializerOptions
                    {
                        WriteIndented = true
                    });

                    // -----------------------------------------------------------------
                    // 5. Output the JSON (to console and to a file)
                    // -----------------------------------------------------------------
                    Console.WriteLine("=== JSON Metadata with Embedded TIFF Thumbnail ===");
                    Console.WriteLine(jsonOutput);

                    string jsonFilePath = "WorkbookMetadata.json";
                    File.WriteAllText(jsonFilePath, jsonOutput, Encoding.UTF8);
                    Console.WriteLine($"\nJSON metadata written to: {Path.GetFullPath(jsonFilePath)}");
                }

                // Optional: Save the original workbook for reference
                string workbookPath = "DemoWorkbook.xlsx";
                workbook.Save(workbookPath, SaveFormat.Xlsx);
                Console.WriteLine($"\nSample workbook saved to: {Path.GetFullPath(workbookPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
