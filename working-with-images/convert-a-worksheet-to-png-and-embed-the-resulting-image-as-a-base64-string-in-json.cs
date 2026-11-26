// Title: Render an Excel worksheet to a PNG image and store it as a Base64 string in a JSON file using Aspose.Cells for .NET
// Common Searches: how to use Aspose.Cells SheetRender to export a worksheet as PNG in C# | C# convert Excel sheet image to Base64 and write to JSON | Aspose.Cells render entire worksheet to single PNG image | save rendered worksheet image as Base64 string in JSON file with .NET | set ImageOrPrintOptions DPI for PNG export in Aspose.Cells
// Tags: Aspose.Cells SheetRender PNG export | C# convert worksheet image to Base64 | ImageOrPrintOptions DPI configuration Aspose.Cells | OnePagePerSheet worksheet rendering | JSON serialization of Base64 image C#

using System;
using System.IO;
using System.Text;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook, renders the first worksheet to a PNG image using Aspose.Cells' SheetRender with custom ImageOrPrintOptions (including DPI and OnePagePerSheet), converts the image bytes to a Base64 string, creates a JSON payload containing the worksheet name, image format, and Base64 data, and writes the formatted JSON to a file.
class Program
{
    static void Main(string[] args)
    {
        // Paths – adjust as needed
        const string inputExcelPath = "input.xlsx";
        const string outputJsonPath = "output.json";

        // Verify input file exists
        if (!File.Exists(inputExcelPath))
        {
            Console.Error.WriteLine($"Error: The input file \"{inputExcelPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputExcelPath);

            // Choose the worksheet to render (first worksheet in this example)
            Worksheet worksheet = workbook.Worksheets[0];

            // Configure rendering options for PNG output
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png,
                OnePagePerSheet = true,          // Render the whole sheet on a single page
                HorizontalResolution = 300,      // DPI settings (optional)
                VerticalResolution = 300
            };

            // Render the worksheet to a memory stream
            using (MemoryStream imageStream = new MemoryStream())
            {
                SheetRender sheetRender = new SheetRender(worksheet, renderOptions);
                // Render the first (and only) page because OnePagePerSheet = true
                sheetRender.ToImage(0, imageStream);

                // Convert the image bytes to a Base64 string
                byte[] imageBytes = imageStream.ToArray();
                string base64Image = Convert.ToBase64String(imageBytes);

                // Build the JSON payload
                var jsonPayload = new
                {
                    WorksheetName = worksheet.Name,
                    ImageFormat = "png",
                    ImageBase64 = base64Image
                };

                // Serialize with indentation for readability
                JsonSerializerOptions jsonOptions = new JsonSerializerOptions
                {
                    WriteIndented = true
                };
                string jsonString = JsonSerializer.Serialize(jsonPayload, jsonOptions);

                // Write JSON to file
                File.WriteAllText(outputJsonPath, jsonString, Encoding.UTF8);
                Console.WriteLine($"JSON successfully written to \"{outputJsonPath}\".");
            }
        }
        catch (Exception ex)
        {
            // Report any unexpected errors without suppressing them
            Console.Error.WriteLine("An error occurred during processing:");
            Console.Error.WriteLine(ex.Message);
            Console.Error.WriteLine(ex.StackTrace);
        }
    }
}
