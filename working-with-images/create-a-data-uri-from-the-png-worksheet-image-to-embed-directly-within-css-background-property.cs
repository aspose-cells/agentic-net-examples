// Title: Generate a PNG data URI from an Excel worksheet using Aspose.Cells for .NET to embed in a CSS background
// AI Prompts: Render the first worksheet of an .xlsx file to a PNG image, convert the image bytes to a Base64 string, and output a data:image/png URI for CSS. | Write C# code that uses SheetRender and ImageOrPrintOptions to produce a memory stream, then builds a background-image rule with the resulting data URI. | Show how to handle file‑not‑found and rendering exceptions while creating a CSS background‑image property from an Excel sheet image.
// Common Searches: how to create a base64 PNG data URI from an Excel sheet with Aspose.Cells in C# | C# Aspose.Cells render worksheet to image and embed in CSS background | convert worksheet image stream to data:image/png for web styling | generate CSS background-image rule from Excel workbook using Aspose.Cells | Aspose.Cells SheetRender ToImage to base64 string example
// Tags: Aspose.Cells render worksheet to PNG | C# convert memory stream to base64 data URI | embed base64 PNG in CSS background-image | SheetRender ToImage with ImageOrPrintOptions | generate data:image/png URI from Excel workbook

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel file, uses Aspose.Cells SheetRender with default ImageOrPrintOptions to render the first worksheet to a PNG image in a memory stream, converts the image bytes to a Base64 string, builds a data:image/png;base64 URI, and prints a CSS background-image rule containing that URI.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Get the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        try
        {
            // Prepare a memory stream for the PNG image
            using (MemoryStream imageStream = new MemoryStream())
            {
                // Render the worksheet to the memory stream (default format is PNG)
                SheetRender sheetRender = new SheetRender(worksheet, new ImageOrPrintOptions());
                sheetRender.ToImage(0, imageStream); // 0 = first page/tab

                // Convert the image bytes to a Base64 string
                string base64 = Convert.ToBase64String(imageStream.ToArray());

                // Build the data URI for PNG
                string dataUri = $"data:image/png;base64,{base64}";

                // Example CSS background property using the data URI
                string cssBackground = $"background-image: url('{dataUri}');";

                // Output the CSS string
                Console.WriteLine(cssBackground);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during rendering or conversion: {ex.Message}");
        }
    }
}
