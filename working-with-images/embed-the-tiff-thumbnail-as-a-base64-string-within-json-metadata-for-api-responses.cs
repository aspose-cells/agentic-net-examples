// Title: Generate JSON metadata with a Base64‑encoded TIFF thumbnail of an Excel workbook using Aspose.Cells in C#
// Common Searches: how to create a Base64 string of a TIFF thumbnail from an Excel file with Aspose.Cells C# | C# Aspose.Cells export workbook to TIFF and embed in JSON response | serialize Excel workbook thumbnail as Base64 in JSON for API using Aspose.Cells | Aspose.Cells generate TIFF image of workbook and convert to Base64 in .NET | return workbook thumbnail as Base64 in JSON from ASP.NET Core API
// Tags: Aspose.Cells export workbook to TIFF | C# convert TIFF to Base64 | JSON metadata with embedded image | Excel workbook thumbnail generation | API response with Base64 encoded image

using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// The WorkbookThumbnailJsonGenerator class loads an Excel workbook with Aspose.Cells, saves it as a TIFF image to a memory stream, converts the TIFF bytes to a Base64 string, and builds a formatted JSON object containing the file name, generation timestamp, thumbnail format, and the Base64‑encoded TIFF. Errors are caught and returned as a JSON error object.
public class WorkbookThumbnailJsonGenerator
{
    // Generates JSON metadata containing a base64‑encoded TIFF representation of the workbook.
    public static string GenerateMetadata(string workbookPath, int thumbnailWidth = 200, int thumbnailHeight = 200)
    {
        try
        {
            // Verify that the workbook file exists to avoid FileNotFoundException.
            if (!File.Exists(workbookPath))
                throw new FileNotFoundException($"Workbook file not found: {workbookPath}");

            // Load the workbook from the specified file.
            Workbook workbook = new Workbook(workbookPath);

            // Save the workbook as a TIFF image to a memory stream.
            // The thumbnail dimensions are not directly used here because Aspose.Cells
            // renders the full sheet(s). If needed, further image processing can be added.
            byte[] tiffBytes;
            using (var ms = new MemoryStream())
            {
                workbook.Save(ms, SaveFormat.Tiff);
                tiffBytes = ms.ToArray();
            }

            // Encode the TIFF byte array as a Base64 string.
            string base64Tiff = Convert.ToBase64String(tiffBytes);

            // Build a JSON object with the desired metadata.
            var metadata = new
            {
                FileName = Path.GetFileName(workbookPath),
                GeneratedOn = DateTime.UtcNow,
                ThumbnailFormat = "tiff",
                ThumbnailBase64 = base64Tiff
            };

            // Serialize the metadata object to a JSON string.
            return JsonSerializer.Serialize(metadata, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (Exception ex)
        {
            // Return a JSON containing the error information.
            var errorInfo = new
            {
                Error = ex.Message,
                StackTrace = ex.StackTrace
            };
            return JsonSerializer.Serialize(errorInfo, new JsonSerializerOptions { WriteIndented = true });
        }
    }

    // Example usage.
    public static void Main()
    {
        string workbookPath = @"C:\Docs\Sample.xlsx";

        // Generate JSON metadata with embedded TIFF representation.
        string jsonMetadata = GenerateMetadata(workbookPath);

        // Output the JSON (e.g., to console or API response).
        Console.WriteLine(jsonMetadata);
    }
}
