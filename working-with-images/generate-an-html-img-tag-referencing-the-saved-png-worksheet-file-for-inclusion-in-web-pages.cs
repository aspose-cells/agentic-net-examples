// Title: Export an Excel worksheet to PNG and output an HTML <img> tag using Aspose.Cells in C#
// AI Prompts: Create C# code that loads an Excel workbook, uses Aspose.Cells SheetRender to save the first worksheet as a PNG file, and writes the resulting <img> element to the console. | Update the sample to set a custom alt attribute and reference the PNG with a relative URL in the generated <img> markup. | Add a verification step that confirms the PNG file was created before emitting the <img> element.
// Common Searches: how to use Aspose.Cells SheetRender in C# to convert an Excel sheet to a PNG image | c# reference a PNG file generated from an Excel worksheet in a web page with Aspose.Cells | check that a PNG file exists after rendering a worksheet before outputting markup in C#
// Tags: png export of worksheet via Aspose.Cells | onepagepersheet option in ImageOrPrintOptions | c# generate img element for exported worksheet image | check file existence after Aspose.Cells rendering | convert excel sheet to image with Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads input.xlsx, renders the first worksheet to Sheet1.png using Aspose.Cells, verifies the PNG file, and prints an HTML <img> tag that references the generated image.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputImage = "Sheet1.png";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (index 0)
            Worksheet sheet = workbook.Worksheets[0];

            // Set up image export options (default format is PNG)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true // Export the whole sheet on one page
            };

            // Create a renderer for the worksheet with the defined options
            SheetRender sheetRender = new SheetRender(sheet, imgOptions);

            // Export the first (and only) page of the worksheet to a PNG file
            sheetRender.ToImage(0, outputImage);

            // Generate an HTML <img> tag that references the saved PNG file
            string imgTag = $"<img src=\"{outputImage}\" alt=\"Worksheet Image\" />";
            Console.WriteLine(imgTag);
        }
        catch (Exception ex)
        {
            // Output any runtime errors for debugging purposes
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
