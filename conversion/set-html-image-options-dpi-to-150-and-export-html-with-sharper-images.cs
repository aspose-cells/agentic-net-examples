// Title: Export Excel to HTML with 150 DPI images using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set image DPI when saving workbook as HTML | How to increase image resolution in HTML output from Aspose.Cells | HtmlSaveOptions ImageOptions HorizontalResolution 150 DPI example | C# export Excel to HTML with high-resolution embedded images | Adjust image DPI for HTML conversion using Aspose.Cells .NET
// Tags: HtmlSaveOptions ImageOptions DPI configuration | Aspose.Cells high-resolution HTML image export | C# set horizontal and vertical resolution for HTML images | Excel to HTML conversion with 150 DPI images | Aspose.Cells image resolution tuning for HTML output

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, configures HtmlSaveOptions.ImageOptions to use a 150 DPI horizontal and vertical resolution for sharper embedded images, and saves the workbook as an HTML file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options: set image DPI to 150 for sharper images
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            htmlOptions.ImageOptions.HorizontalResolution = 150;
            htmlOptions.ImageOptions.VerticalResolution = 150;

            // Export the workbook to HTML using the configured options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
