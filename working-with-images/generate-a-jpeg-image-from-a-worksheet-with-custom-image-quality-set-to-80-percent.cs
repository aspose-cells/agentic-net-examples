// Title: Create a JPEG image from an Excel worksheet with 80% quality using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export worksheet to JPEG with specific quality setting | How to set JPEG compression level when rendering Excel sheet to image in .NET | Aspose.Cells ImageOrPrintOptions Quality property example for JPEG output | Render Excel worksheet as high‑resolution JPEG using SheetRender in C# | Save Excel sheet as JPEG with 80 percent quality using Aspose.Cells API
// Tags: ImageOrPrintOptions JPEG quality configuration | SheetRender ToImage method example | Aspose.Cells JPEG rendering options | C# Excel worksheet image export workflow | OnePagePerSheet rendering flag

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace WorksheetToJpegExample
{
    // The example loads (or creates) an Excel workbook, configures ImageOrPrintOptions to output JPEG with Quality=80, sets OnePagePerSheet and 150 dpi resolution, and uses SheetRender to render the first worksheet to a JPEG file.
    class Program
    {
        static void Main(string[] args)
        {
            // Define paths for the input workbook and the output JPEG image
            string inputFile = "SampleData.xlsx";
            string outputFile = "WorksheetImage.jpg";

            try
            {
                Workbook workbook;

                // Try to load an existing workbook; if it does not exist, create one with sample data
                if (File.Exists(inputFile))
                {
                    workbook = new Workbook(inputFile);
                    Console.WriteLine($"Loaded workbook from \"{inputFile}\".");
                }
                else
                {
                    workbook = new Workbook();
                    Worksheet ws = workbook.Worksheets[0];

                    // Populate the worksheet with some sample data
                    ws.Cells["A1"].PutValue("Product");
                    ws.Cells["B1"].PutValue("Quantity");
                    ws.Cells["A2"].PutValue("Apple");
                    ws.Cells["B2"].PutValue(120);
                    ws.Cells["A3"].PutValue("Orange");
                    ws.Cells["B3"].PutValue(85);
                    ws.Cells["A4"].PutValue("Banana");
                    ws.Cells["B4"].PutValue(150);

                    // Save the sample workbook for future runs
                    workbook.Save(inputFile);
                    Console.WriteLine($"Created sample workbook \"{inputFile}\" with test data.");
                }

                // Access the first worksheet (or any specific worksheet you need)
                Worksheet worksheet = workbook.Worksheets[0];

                // Configure image rendering options
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
                {
                    // Set the output format to JPEG
                    ImageType = ImageType.Jpeg,

                    // Set JPEG quality to 80 percent (value range is 0‑100)
                    Quality = 80,

                    // Render the whole sheet on a single page (optional, but often desired for images)
                    OnePagePerSheet = true,

                    // Optional: increase resolution for sharper output
                    HorizontalResolution = 150,
                    VerticalResolution = 150
                };

                // Create a SheetRender object with the worksheet and rendering options
                SheetRender sheetRender = new SheetRender(worksheet, renderOptions);

                // Render the first (and only) page to a JPEG file
                sheetRender.ToImage(0, outputFile);

                Console.WriteLine($"Worksheet successfully rendered to JPEG image \"{outputFile}\" with 80% quality.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during the rendering process:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
