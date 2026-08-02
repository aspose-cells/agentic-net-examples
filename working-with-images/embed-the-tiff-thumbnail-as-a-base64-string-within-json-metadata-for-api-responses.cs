// Title: C# – Generate JSON with Base64‑encoded TIFF thumbnail of an Excel worksheet using Aspose.Cells
// Description: Loads an Excel workbook, renders the first worksheet to a single‑page TIFF at 150 dpi with Aspose.Cells SheetRender, converts the image to a Base64 string, and returns an indented JSON payload that includes the file name, thumbnail format, Base64 data and a UTC timestamp. Perfect for API responses or metadata storage.
// Keywords: Aspose.Cells | C# | TIFF thumbnail | Base64 JSON | SheetRender | ToTiff | Excel preview | API payload | image serialization | .NET
// Common Searches: Aspose.Cells generate TIFF thumbnail C# | Base64 image in JSON using Aspose.Cells | Excel worksheet preview API C# | How to embed Excel thumbnail in JSON response | C# convert TIFF to Base64 string
// Developer Intent: Create a JSON response that embeds a Base64‑encoded TIFF preview of the first worksheet in an Excel file.
// Use Cases: Return the thumbnail in a REST API so client apps can show a quick preview without downloading the full workbook. | Store the Base64 thumbnail with file metadata in a database for searchable document catalogs. | Send the preview to a web UI for instant display in a file‑manager while the Excel file loads asynchronously.
// AI Prompts: Write a C# method that accepts an Excel file path and returns a JSON string containing a Base64‑encoded TIFF thumbnail of the first worksheet using Aspose.Cells. | Extend the sample to let the caller specify the worksheet index and custom DPI values for the generated thumbnail. | Add error handling that returns a structured JSON error object when the source file is missing or thumbnail generation fails.

using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsThumbnailJson
{
    // Loads an Excel workbook, renders the first worksheet to a single‑page TIFF at 150 dpi with Aspose.Cells SheetRender, converts the image to a Base64 string, and returns an indented JSON payload that includes the file name, thumbnail format, Base64 data and a UTC timestamp. Perfect for API responses or metadata storage.
    public class Program
    {
        // Entry point
        public static void Main(string[] args)
        {
            try
            {
                // Path to the source Excel file
                string excelPath = "sample.xlsx";

                // Verify that the source file exists to avoid FileNotFoundException
                if (!File.Exists(excelPath))
                {
                    Console.Error.WriteLine($"Error: The file '{excelPath}' was not found.");
                    return;
                }

                // Generate JSON containing the TIFF thumbnail as a Base64 string
                string jsonResult = GenerateThumbnailJson(excelPath);

                // Output the JSON (for demonstration purposes)
                Console.WriteLine(jsonResult);
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Generates a JSON string with a Base64‑encoded TIFF thumbnail of the first worksheet
        private static string GenerateThumbnailJson(string workbookPath)
        {
            try
            {
                // Load the workbook (lifecycle rule: load)
                Workbook workbook = new Workbook(workbookPath);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Configure rendering options for a single‑page TIFF thumbnail
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
                {
                    OnePagePerSheet = true,          // Render the whole sheet on one page
                    HorizontalResolution = 150,     // Reasonable resolution for a thumbnail
                    VerticalResolution = 150
                    // ImageFormat is not required for ToTiff rendering
                };

                // Create a SheetRender instance (lifecycle rule: create)
                SheetRender renderer = new SheetRender(sheet, renderOptions);

                // Render the worksheet to a memory stream as TIFF (lifecycle rule: save to stream)
                using (MemoryStream tiffStream = new MemoryStream())
                {
                    renderer.ToTiff(tiffStream); // Render to TIFF

                    // Ensure the stream position is at the beginning
                    tiffStream.Position = 0;

                    // Convert the TIFF bytes to a Base64 string
                    string base64Thumbnail = Convert.ToBase64String(tiffStream.ToArray());

                    // Build an anonymous object for JSON serialization
                    var metadata = new
                    {
                        FileName = Path.GetFileName(workbookPath),
                        ThumbnailFormat = "tiff",
                        ThumbnailBase64 = base64Thumbnail,
                        GeneratedOn = DateTime.UtcNow.ToString("o")
                    };

                    // Serialize the object to JSON
                    JsonSerializerOptions jsonOptions = new JsonSerializerOptions
                    {
                        WriteIndented = true
                    };
                    return JsonSerializer.Serialize(metadata, jsonOptions);
                }
            }
            catch (Exception ex)
            {
                // Log any errors that occur during thumbnail generation
                Console.Error.WriteLine($"Failed to generate thumbnail JSON: {ex.Message}");
                return string.Empty;
            }
        }
    }
}
