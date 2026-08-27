// Title: Generate HTML from an Aspose.Cells workbook with 200 DPI external images using C#
// AI Prompts: Write C# code that creates a workbook, inserts a picture, and saves it as HTML with 200 DPI images using Aspose.Cells. | Show how to configure HtmlSaveOptions.ImageOptions to set HorizontalResolution and VerticalResolution to 200 DPI and export images as separate files. | Demonstrate disabling Base64 image embedding while exporting a workbook to HTML for high‑resolution image output in C#.
// Common Searches: Aspose.Cells C# export workbook to HTML with high resolution images | How to set image DPI to 200 when saving Excel as HTML using Aspose.Cells | Export Excel pictures as external files instead of Base64 with Aspose.Cells | HtmlSaveOptions ImageOptions HorizontalResolution 200 DPI example | C# generate HTML from workbook with external PNG images at 200 DPI
// Tags: Aspose.Cells HtmlSaveOptions image DPI | C# export workbook to HTML external images | Aspose.Cells set HorizontalResolution verticalResolution | high‑resolution image export Aspose.Cells | disable Base64 image embedding Aspose.Cells HTML

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample creates a workbook, optionally adds a picture, configures HtmlSaveOptions.ImageOptions to 200 DPI for both horizontal and vertical resolution, disables Base64 image embedding, and saves the workbook as an HTML file with high‑resolution external image files.
class ExportHtmlHighResolution
{
    static void Main()
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook and add some sample content
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("High‑Resolution HTML Export");

        // Add an image to the worksheet if the file exists
        string imagePath = "example.png";
        if (File.Exists(imagePath))
        {
            // Add picture at row index 2 (third row), column index 0 (first column)
            sheet.Pictures.Add(2, 0, imagePath);
        }
        else
        {
            Console.WriteLine($"Image file not found: {imagePath}. Skipping picture insertion.");
        }

        // Configure HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();

        // Set high‑resolution image options
        ImageOrPrintOptions imageOpts = saveOptions.ImageOptions;
        imageOpts.HorizontalResolution = 200;
        imageOpts.VerticalResolution = 200;

        // Export images as separate files (not Base64) to preserve high resolution
        saveOptions.ExportImagesAsBase64 = false;

        // Save the workbook as an HTML file with the specified high‑resolution image settings
        string outputPath = "high_res_output.html";
        workbook.Save(outputPath, saveOptions);
        Console.WriteLine($"HTML file saved to: {Path.GetFullPath(outputPath)}");
    }
}
