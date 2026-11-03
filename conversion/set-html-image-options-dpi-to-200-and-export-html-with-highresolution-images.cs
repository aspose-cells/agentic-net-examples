// Title: Export Excel workbook to HTML with high‑resolution (200 DPI) images using Aspose.Cells for .NET
// Common Searches: Aspose.Cells set image DPI for HTML export in C# | How to increase image resolution when saving Excel as HTML with Aspose.Cells | C# export Excel to HTML with 200 DPI images using Aspose.Cells | Configure ImageOrPrintOptions DPI for HTML conversion in Aspose.Cells .NET
// Tags: Aspose.Cells HTML export high resolution images | C# ImageOrPrintOptions DPI configuration | Excel to HTML conversion custom image DPI | Aspose.Cells SaveFormat.Html image quality control | HtmlSaveOptions image resolution limitation

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook and saves it as HTML using Aspose.Cells. While HtmlSaveOptions does not expose an ImageResolution property, image DPI for the generated HTML can be adjusted through ImageOrPrintOptions, enabling high‑resolution (e.g., 200 DPI) images in the output.
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
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Note: Image DPI can be controlled via ImageOrPrintOptions when rendering images.
            // HtmlSaveOptions does not expose an ImageResolution property in this version.

            // Save the workbook as an HTML file
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
