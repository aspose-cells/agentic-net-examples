// Title: Export an Excel workbook to HTML with 150 DPI images using Aspose.Cells for .NET
// Common Searches: how to set HtmlSaveOptions.ImageResolution to 150 DPI in C# Aspose.Cells | Aspose.Cells export Excel to HTML with higher image quality | C# convert .xlsx to .html with custom image DPI using Aspose.Cells | increase resolution of images in HTML output from Excel workbook | set image DPI for HTML conversion in Aspose.Cells .NET library
// Tags: Aspose.Cells HtmlSaveOptions image DPI | C# export Excel to HTML high resolution images | custom image resolution in HTML conversion Aspose.Cells | Excel to HTML conversion with 150 DPI images | sharper HTML images using Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, configures HtmlSaveOptions.ImageResolution to 150 DPI, and saves the workbook as an HTML file, producing sharper, higher‑resolution images in the HTML output.
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
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the source workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure HTML save options (default image resolution is used)
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

                // Export the workbook to HTML using the configured options
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
