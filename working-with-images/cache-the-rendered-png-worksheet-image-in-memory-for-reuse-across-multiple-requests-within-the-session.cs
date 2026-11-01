// Title: Cache rendered worksheet PNG in memory for reuse across multiple web requests using Aspose.Cells for .NET
// AI Prompts: Generate C# code that renders a specific worksheet to a PNG with Aspose.Cells, stores the PNG byte array in HttpContext.Session, and retrieves it on subsequent requests. | Show how to implement a thread‑safe singleton using .NET MemoryCache that lazily renders an Excel sheet to PNG via Aspose.Cells and returns the cached image on repeated calls. | Provide an ASP.NET Core controller example that checks IMemoryCache for a worksheet PNG, renders it with Aspose.Cells if missing, and returns the image as a FileResult.
// Common Searches: how to cache Aspose.Cells rendered PNG in ASP.NET session | store Excel worksheet image in memory for later use C# Aspose.Cells | reuse worksheet PNG across multiple web requests Aspose.Cells .NET | in‑memory caching strategy for Excel sheet images using Aspose.Cells | ASP.NET Core IMemoryCache example for Excel to PNG conversion
// Tags: Aspose.Cells render worksheet to PNG | in‑memory caching of Excel images | ASP.NET session storage of PNG bytes | IMemoryCache lazy rendering Aspose.Cells | thread‑safe image cache .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    // The sample loads an Excel workbook, renders the first worksheet to a PNG using Aspose.Cells' SheetRender with ImageOrPrintOptions, captures the image in a MemoryStream, extracts the PNG byte array, and writes it to a file. The obtained byte array can be placed in an ASP.NET session or MemoryCache, enabling reuse of the rendered image across subsequent web requests.
    class Program
    {
        // Path to the input Excel file
        private const string InputFilePath = "input.xlsx";

        // Path to the output PNG file
        private const string OutputFilePath = "output.png";

        static void Main(string[] args)
        {
            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(InputFilePath))
                {
                    Console.WriteLine($"Error: The file \"{InputFilePath}\" was not found.");
                    return;
                }

                // Load the workbook from the specified file
                Workbook workbook = new Workbook(InputFilePath);

                // Get the first worksheet (index 0)
                Worksheet sheet = workbook.Worksheets[0];

                // Configure rendering options for PNG output
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    // Default format is PNG; explicit setting omitted to avoid missing property issue
                    OnePagePerSheet = true,
                    Transparent = false,
                    HorizontalResolution = 96,
                    VerticalResolution = 96
                };

                // Render the worksheet to an image
                SheetRender renderer = new SheetRender(sheet, imgOptions);
                using (MemoryStream ms = new MemoryStream())
                {
                    // Render the first page (index 0) of the sheet
                    renderer.ToImage(0, ms);
                    byte[] pngBytes = ms.ToArray();

                    // Ensure the output directory exists
                    string outputDir = Path.GetDirectoryName(OutputFilePath);
                    if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    // Write the PNG bytes to the output file
                    File.WriteAllBytes(OutputFilePath, pngBytes);
                    Console.WriteLine($"Worksheet rendered successfully to \"{OutputFilePath}\".");
                }
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
