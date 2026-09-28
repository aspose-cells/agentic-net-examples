// Title: C# – Create a base64‑encoded TIFF thumbnail from an Excel workbook and embed it in JSON metadata with Aspose.Cells
// AI Prompts: Write C# code that loads an .xlsx file using Aspose.Cells, converts the workbook to a TIFF image in a MemoryStream, encodes the image to a base64 string, and returns a JSON object containing the file name, generation timestamp, format, and the base64 thumbnail. | Extend the program to accept a command‑line argument specifying the thumbnail format (TIFF, PNG, or JPEG) and update the JSON metadata with the selected format and corresponding base64 image. | Add robust error handling that catches missing input files or conversion failures and outputs a JSON error object with a message, error code, and optional stack trace.
// Common Searches: aspnet core return excel thumbnail as base64 json using aspose.cells | c# convert excel workbook to tiff in memory and embed in json response | how to generate excel file thumbnail for api output with aspose.cells | serialize tiff image to base64 string in c# for json metadata
// Tags: Aspose.Cells convert workbook to TIFF in memory | C# base64 encode TIFF thumbnail | JSON metadata with embedded image data | Excel file thumbnail generation for API | SaveFormat.Tiff usage in Aspose.Cells | MemoryStream image conversion C#

using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// The sample loads an Excel workbook, saves it as a TIFF image directly into a MemoryStream using Aspose.Cells, converts the TIFF bytes to a base64 string, and builds a formatted JSON object that includes the original file name, UTC generation timestamp, thumbnail format, and the base64‑encoded thumbnail.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.Error.WriteLine($"Error: File '{inputPath}' not found.");
            return;
        }

        try
        {
            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Convert the workbook to a TIFF image in memory (acts as a thumbnail)
            string base64Tiff;
            using (var ms = new MemoryStream())
            {
                // Save as TIFF; you can adjust the page count or scaling if needed
                workbook.Save(ms, SaveFormat.Tiff);
                byte[] tiffBytes = ms.ToArray();
                base64Tiff = Convert.ToBase64String(tiffBytes);
            }

            // Build JSON metadata containing the base64 TIFF thumbnail
            var metadata = new
            {
                FileName = Path.GetFileName(inputPath),
                GeneratedOn = DateTime.UtcNow,
                ThumbnailFormat = "tiff",
                ThumbnailBase64 = base64Tiff
            };

            // Serialize to JSON string with indentation
            string json = JsonSerializer.Serialize(metadata, new JsonSerializerOptions { WriteIndented = true });

            // Output JSON
            Console.WriteLine(json);
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and output an error message
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
