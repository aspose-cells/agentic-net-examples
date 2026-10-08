// Title: Render an Excel worksheet to a PNG image and embed it as a Base64 string in a JSON payload using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file, uses Aspose.Cells SheetRender with ImageOrPrintOptions to generate a PNG of the first worksheet at 150 dpi, converts the PNG bytes to a Base64 string, and returns a JSON object containing the image data. | Create a .NET console application that verifies the existence of an Excel workbook, renders the selected worksheet to a single‑page PNG stream using Aspose.Cells, encodes the stream to Base64, and serializes the result with System.Text.Json. | Provide a C# snippet that demonstrates how to set OnePagePerSheet, render a worksheet to a MemoryStream as PNG, and embed the Base64‑encoded image into a JSON structure for API consumption.
// Common Searches: c# aspocells convert worksheet to png and get base64 string | how to render excel sheet as png and embed in json using .net | aspocells export single worksheet to png with specific dpi | serialize png image from excel to json in c# console app | base64 encode excel worksheet image for web api
// Tags: aspocells render worksheet to png | c# convert excel sheet to base64 image | json serialize base64 png in .net | imageorprintoptions set dpi aspocells | sheetrender single page png export

using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The program loads an Excel file, renders the first worksheet to a PNG image at 150 dpi using Aspose.Cells, converts the image bytes to a Base64 string, and outputs a JSON object that contains the encoded image.
class WorksheetToPngJson
{
    static void Main()
    {
        try
        {
            // Path to the source Excel file
            string excelPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(excelPath))
            {
                Console.Error.WriteLine($"Error: The file '{excelPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Choose the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set image rendering options
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // Optional: set the resolution (dpi)
                HorizontalResolution = 150,
                VerticalResolution = 150,
                // Fit the whole sheet on one page
                OnePagePerSheet = true
                // Note: ImageFormat defaults to PNG, so no explicit setting is required
            };

            // Render the worksheet to an image
            SheetRender sheetRender = new SheetRender(sheet, imgOptions);
            using (MemoryStream pngStream = new MemoryStream())
            {
                // Render the first page (index 0)
                sheetRender.ToImage(0, pngStream);
                pngStream.Position = 0; // Reset stream position

                // Convert PNG bytes to Base64 string
                string base64Image = Convert.ToBase64String(pngStream.ToArray());

                // Build JSON object containing the Base64 image string
                var jsonObj = new
                {
                    image = base64Image
                };

                // Serialize to JSON (indented for readability)
                string json = JsonSerializer.Serialize(jsonObj, new JsonSerializerOptions { WriteIndented = true });

                // Output the JSON
                Console.WriteLine(json);
            }
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
