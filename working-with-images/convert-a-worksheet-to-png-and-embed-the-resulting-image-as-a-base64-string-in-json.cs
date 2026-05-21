using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Configure rendering options for PNG output
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Png,
            OnePagePerSheet = true
        };

        // Create a SheetRender instance for the worksheet
        SheetRender sheetRender = new SheetRender(worksheet, options);

        // Render the first page of the worksheet to a memory stream
        using (MemoryStream imageStream = new MemoryStream())
        {
            sheetRender.ToImage(0, imageStream); // Uses SheetRender.ToImage(int, Stream) rule
            byte[] imageBytes = imageStream.ToArray();

            // Convert the image bytes to a Base64 string
            string base64Image = Convert.ToBase64String(imageBytes);

            // Create a simple JSON object containing the Base64 image
            var jsonObject = new { imageBase64 = base64Image };
            string json = JsonSerializer.Serialize(jsonObject);

            // Output the JSON string
            Console.WriteLine(json);
        }

        // Clean up resources
        sheetRender.Dispose();
    }
}