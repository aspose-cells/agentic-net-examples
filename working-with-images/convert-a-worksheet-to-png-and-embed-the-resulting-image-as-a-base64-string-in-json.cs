// Title: Render a named Excel worksheet to PNG and embed the image as a Base64 string in JSON using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, uses Aspose.Cells SheetRender with ImageOrPrintOptions to create a PNG stream of a specific worksheet, converts the stream to a Base64 string, and serializes it into a JSON object. | Show how to verify the Excel file exists, render the first page of a worksheet to a MemoryStream as PNG, then use Convert.ToBase64String and System.Text.Json to output the image data in JSON. | Provide a console application example that selects a worksheet by name, sets OnePagePerSheet = true, renders it to PNG, encodes the bytes to Base64, and prints the resulting JSON payload.
// Common Searches: c# aspnet convert excel sheet to png and return base64 json | how to use Aspose.Cells SheetRender to get a PNG stream from a worksheet | serialize png image from Excel as base64 in .NET console app | export specific worksheet as png and embed in json using Aspose.Cells
// Tags: Aspose.Cells SheetRender PNG export | C# convert worksheet to base64 image | ImageOrPrintOptions OnePagePerSheet usage | System.Text.Json serialize base64 image | MemoryStream PNG rendering Aspose.Cells

using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads 'input.xlsx', renders the worksheet named 'Sheet1' to a PNG image using Aspose.Cells SheetRender with OnePagePerSheet enabled, converts the PNG bytes to a Base64 string, wraps it in a JSON object via System.Text.Json, and writes the JSON to the console.
class Program
{
    static void Main(string[] args)
    {
        // Path to the source Excel file
        string excelPath = "input.xlsx";

        // Name of the worksheet to convert
        string worksheetName = "Sheet1";

        // Verify that the input file exists
        if (!File.Exists(excelPath))
        {
            Console.Error.WriteLine($"File not found: {excelPath}");
            return;
        }

        try
        {
            // Load the workbook from file
            Workbook workbook = new Workbook(excelPath);

            // Retrieve the worksheet by name
            Worksheet worksheet = workbook.Worksheets[worksheetName];
            if (worksheet == null)
            {
                Console.Error.WriteLine($"Worksheet '{worksheetName}' not found.");
                return;
            }

            // Set image rendering options (default PNG, one page per sheet)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true
            };

            // Create a SheetRender object for the worksheet
            SheetRender sheetRender = new SheetRender(worksheet, imgOptions);

            // Render the worksheet to a memory stream as PNG
            using (MemoryStream pngStream = new MemoryStream())
            {
                // Render the first (and only) page of the sheet
                sheetRender.ToImage(0, pngStream);

                // Convert the PNG bytes to a Base64 string
                string base64Image = Convert.ToBase64String(pngStream.ToArray());

                // Build a simple JSON object containing the Base64 image
                var jsonObject = new { image = base64Image };
                string json = JsonSerializer.Serialize(jsonObject);

                // Output the JSON string
                Console.WriteLine(json);
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
