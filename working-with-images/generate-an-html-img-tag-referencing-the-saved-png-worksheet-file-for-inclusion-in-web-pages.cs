// Title: Create a relative HTML <img> tag for a PNG worksheet image generated with Aspose.Cells in C#
// Common Searches: C# Aspose.Cells render first worksheet to PNG and get HTML img tag | How to embed Excel worksheet image in a web page using Aspose.Cells C# | Generate HTML <img> element for worksheet.png created by Aspose.Cells | Aspose.Cells SheetRender ToImage output PNG and build img tag in C# | C# code to convert Excel sheet to PNG and produce <img> markup
// Tags: Aspose.Cells SheetRender ToImage PNG | ImageOrPrintOptions OnePagePerSheet C# | render Excel worksheet to PNG Aspose.Cells | build HTML img tag from file path C# | embed worksheet image in web page Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsImageToHtml
{
    // The example loads an Excel workbook, uses Aspose.Cells ImageOrPrintOptions and SheetRender to render the first worksheet as a high‑resolution PNG, then constructs a relative HTML <img> tag that references the saved image file and outputs the markup.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths – adjust as needed or pass via command‑line arguments
            string workbookPath = "input.xlsx";
            string imagePath = "worksheet.png";

            // Validate input workbook existence
            if (!File.Exists(workbookPath))
            {
                Console.Error.WriteLine($"Error: Workbook file \"{workbookPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Get the first worksheet (you can change the index or name as required)
                Worksheet sheet = workbook.Worksheets[0];

                // Configure rendering options for PNG output
                ImageOrPrintOptions options = new ImageOrPrintOptions
                {
                    ImageType = Aspose.Cells.Drawing.ImageType.Png,
                    // Render the whole sheet on a single page – useful for a compact image
                    OnePagePerSheet = true,
                    // Optional: set resolution for higher quality
                    HorizontalResolution = 300,
                    VerticalResolution = 300
                };

                // Create the SheetRender object
                SheetRender renderer = new SheetRender(sheet, options);

                // Render the first (and only) page to a PNG file
                // Page index is zero‑based
                renderer.ToImage(0, imagePath);

                // Dispose the renderer (optional, as it implements IDisposable)
                renderer.Dispose();

                // Build the HTML <img> tag referencing the saved PNG file
                // Use a relative URL assuming the image will be served from the same directory as the HTML page
                string imgTag = $"<img src=\"{Path.GetFileName(imagePath)}\" alt=\"Worksheet Image\" />";

                // Output the HTML tag to the console (or write to a file if desired)
                Console.WriteLine("HTML <img> tag:");
                Console.WriteLine(imgTag);
            }
            catch (Exception ex)
            {
                // Report any unexpected errors
                Console.Error.WriteLine("An error occurred while processing the workbook:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
