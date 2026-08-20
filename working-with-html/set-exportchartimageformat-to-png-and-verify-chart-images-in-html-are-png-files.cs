// Title: Export Aspose.Cells Chart Images as PNG in HTML and Verify Output (C# .NET)
// Description: Creates a workbook with a column chart, sets HtmlSaveOptions.ImageOptions.ImageType to PNG, disables Base64 encoding, saves the workbook as HTML, then scans the generated HTML for <img> tags ending with .png and confirms the corresponding PNG files exist in the output folder.
// Keywords: Aspose.Cells PNG chart export | HtmlSaveOptions ImageType PNG | C# export Excel chart to HTML | verify chart image format Aspose | .NET HTML report PNG images | Aspose.Cells image verification script
// Common Searches: Aspose.Cells export chart as PNG HTML | Set ExportChartImageFormat to PNG Aspose | Check PNG images in Aspose.Cells HTML output | C# verify chart image files after HTML export | Aspose.Cells HtmlSaveOptions ImageOptions example
// Developer Intent: Configure Aspose.Cells to save chart images as PNG files during HTML export and programmatically confirm that the HTML references only PNG images and that the PNG files are present.
// Use Cases: Generate web‑ready HTML reports where all chart graphics are separate PNG files for better browser compatibility. | Automate quality checks in a CI pipeline to ensure exported chart images meet PNG standards before publishing. | Create a batch conversion tool that transforms multiple Excel workbooks to HTML with PNG charts and validates the output files.
// AI Prompts: Write C# code using Aspose.Cells to export a workbook to HTML with chart images saved as PNG files and include a verification step that checks the HTML and file system. | Explain how HtmlSaveOptions.ImageOptions.ImageType affects chart image formats in Aspose.Cells and how to validate the generated HTML for PNG references. | Suggest a way to assert that all chart images in an Aspose.Cells HTML export are PNG without reading the file system, using only HTML content analysis.

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// Creates a workbook with a column chart, sets HtmlSaveOptions.ImageOptions.ImageType to PNG, disables Base64 encoding, saves the workbook as HTML, then scans the generated HTML for <img> tags ending with .png and confirms the corresponding PNG files exist in the output folder.
class ExportChartImageAsPng
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["A4"].PutValue("Banana");

        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(80);
        sheet.Cells["B4"].PutValue(150);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Prepare HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Ensure images are saved as separate files (not Base64)
        htmlOptions.ExportImagesAsBase64 = false;

        // Set the image type for charts and other images to PNG
        // This controls the ExportChartImageFormat behavior
        htmlOptions.ImageOptions.ImageType = ImageType.Png;

        // Define output folder and HTML file path
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "HtmlOutput");
        Directory.CreateDirectory(outputFolder);
        string htmlPath = Path.Combine(outputFolder, "Workbook.html");

        // Save the workbook as HTML
        workbook.Save(htmlPath, htmlOptions);
        Console.WriteLine($"Workbook saved as HTML to: {htmlPath}");

        // ---------- Verification ----------
        // 1. Read the generated HTML and look for image sources ending with .png
        string htmlContent = File.ReadAllText(htmlPath);
        var imgSrcMatches = Regex.Matches(htmlContent, @"<img\s+[^>]*src\s*=\s*[""']([^""']+)[""']", RegexOptions.IgnoreCase);

        bool allPng = true;
        foreach (Match match in imgSrcMatches)
        {
            string src = match.Groups[1].Value;
            if (!src.EndsWith(".png", StringComparison.OrdinalIgnoreCase))
            {
                allPng = false;
                Console.WriteLine($"Non‑PNG image found in HTML: {src}");
            }
        }

        // 2. Verify that the image files with .png extension actually exist in the output folder
        string[] pngFiles = Directory.GetFiles(outputFolder, "*.png");
        if (pngFiles.Length == 0)
        {
            allPng = false;
            Console.WriteLine("No PNG image files were generated.");
        }

        // Report verification result
        if (allPng)
        {
            Console.WriteLine("Verification succeeded: all chart images in the HTML are PNG files.");
        }
        else
        {
            Console.WriteLine("Verification failed: some chart images are not PNG.");
        }
    }
}
