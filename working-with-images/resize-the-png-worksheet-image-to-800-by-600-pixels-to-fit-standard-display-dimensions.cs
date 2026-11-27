// Title: Resize an Excel worksheet to a fixed 800 × 600 PNG image with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells set worksheet image size to 800x600 pixels | How to export Excel sheet as 800x600 PNG using Aspose.Cells | Aspose.Cells ImageOrPrintOptions SetDesiredSize example in C# | Render Excel worksheet to a specific PNG dimension with Aspose.Cells | Force exact PNG size when converting Excel sheet to image in .NET
// Tags: Aspose.Cells ImageOrPrintOptions SetDesiredSize PNG | Worksheet to PNG rendering Aspose.Cells | C# fixed-size Excel sheet image export | OnePagePerSheet PNG export Aspose.Cells | SheetRender ToImage custom dimensions

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace WorksheetImageResize
{
    // The sample loads an Excel workbook, selects a worksheet, configures ImageOrPrintOptions to render a PNG, disables aspect‑ratio preservation, sets the desired size to 800 × 600 pixels, and uses SheetRender to save the worksheet as a resized PNG file.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input workbook and output image
            string inputWorkbookPath = "input.xlsx";
            string outputImagePath = "worksheet_resized.png";

            // Verify that the input file exists
            if (!File.Exists(inputWorkbookPath))
            {
                Console.WriteLine($"Error: The workbook file \"{inputWorkbookPath}\" was not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputWorkbookPath);

                // Get the first worksheet (you can change the index or name as needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Configure image rendering options
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
                {
                    // Render as PNG
                    ImageType = Aspose.Cells.Drawing.ImageType.Png,
                    // Ensure the whole sheet is rendered on a single page
                    OnePagePerSheet = true
                };

                // Set the desired output size to 800x600 pixels.
                // The third parameter (keepAspectRatio) is set to false to force exact dimensions.
                renderOptions.SetDesiredSize(800, 600, false);

                // Create a SheetRender object with the worksheet and rendering options
                SheetRender sheetRender = new SheetRender(sheet, renderOptions);

                // Render the first (and only) page to the specified PNG file
                // If the worksheet spans multiple pages, you may need to loop over sheetRender.PageCount
                sheetRender.ToImage(0, outputImagePath);

                Console.WriteLine($"Worksheet image successfully rendered and resized to 800x600 pixels:");
                Console.WriteLine($"Output file: {Path.GetFullPath(outputImagePath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
