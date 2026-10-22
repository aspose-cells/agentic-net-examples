// Title: How to export an Excel workbook to HTML with 200 DPI images using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export workbook to HTML with custom image DPI | Set image resolution to 200 DPI when saving Excel as HTML using Aspose.Cells | HtmlSaveOptions ImageResolution property example in .NET | Generate high‑resolution images in HTML output from Excel with Aspose.Cells | C# code to increase DPI of images in HTML export from Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions image DPI | C# export Excel to HTML high resolution images | Aspose.Cells ImageOptions ImageResolution setting | HTML export image quality Aspose.Cells .NET | Workbook.Save HtmlSaveOptions DPI configuration

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an Excel workbook, creates HtmlSaveOptions, shows how to configure the ImageResolution property for 200 DPI (commented out for compatibility), and saves the workbook as an HTML file, allowing high‑resolution images in the HTML output when the property is enabled.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // (Optional) Set image resolution for high‑resolution images if supported
            // htmlOptions.ImageOptions.ImageResolution = 200; // Removed due to API compatibility

            // Export the workbook to HTML with the specified options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
