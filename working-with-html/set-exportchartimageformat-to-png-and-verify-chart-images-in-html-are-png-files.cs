// Title: Export Excel Chart as PNG in HTML using Aspose.Cells for .NET and Validate Image Files
// Description: C# example that builds a workbook, adds a column chart, sets HtmlSaveOptions to render chart images as separate PNG files (ImageOptions.ImageType = Png, ExportImagesAsBase64 = false), saves the workbook as HTML, and scans the resulting HTML to confirm every <img> tag referencing a local file ends with .png.
// Keywords: Aspose.Cells | C# | .NET | HTML export | chart PNG | ImageOptions.ImageType | ExportImagesAsBase64 | chart image verification | Excel to HTML | chart rendering
// Common Searches: Aspose.Cells export chart PNG | HTML save options chart image type | verify chart image extension Aspose | C# export Excel chart as PNG HTML | set ImageOptions.ImageType in Aspose.Cells
// Developer Intent: Configure Aspose.Cells to output chart images as PNG files when saving a workbook to HTML and programmatically ensure the HTML references only PNG images.
// Use Cases: Create web‑ready reports where charts are stored as individual PNG files for caching or further processing. | Automate quality checks that guarantee all chart images in generated HTML have the .png extension. | Integrate chart image export into CI pipelines to validate output consistency across environments. | Separate chart assets from HTML for easier localization or theming.
// AI Prompts: Generate C# code that uses Aspose.Cells to save a workbook as HTML with chart images saved as PNG files and then checks the HTML for correct file extensions. | Write a method that parses an HTML file produced by Aspose.Cells and returns true only if every local <img> source ends with .png. | Explain the impact of HtmlSaveOptions.ImageOptions.ImageType and ExportImagesAsBase64 on chart image output in Aspose.Cells. | Provide a step‑by‑step guide to verify chart image formats in HTML generated from an Excel workbook using Aspose.Cells.

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// C# example that builds a workbook, adds a column chart, sets HtmlSaveOptions to render chart images as separate PNG files (ImageOptions.ImageType = Png, ExportImagesAsBase64 = false), saves the workbook as HTML, and scans the resulting HTML to confirm every <img> tag referencing a local file ends with .png.
class ExportChartImageAsPng
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
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

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure HTML save options to export chart images as PNG
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportImagesAsBase64 = false
            };
            // Set the image type for chart rendering to PNG
            htmlOptions.ImageOptions.ImageType = ImageType.Png;

            // Define output paths
            string outputFolder = Path.Combine(Environment.CurrentDirectory, "HtmlOutput");
            Directory.CreateDirectory(outputFolder);
            string htmlPath = Path.Combine(outputFolder, "ChartExport.html");

            // Save the workbook as HTML
            workbook.Save(htmlPath, htmlOptions);
            Console.WriteLine($"Workbook saved as HTML to: {htmlPath}");

            // Verify that chart images referenced in the HTML are PNG files
            if (File.Exists(htmlPath))
            {
                string htmlContent = File.ReadAllText(htmlPath);
                var imgSrcMatches = Regex.Matches(htmlContent, @"<img[^>]+src\s*=\s*[""']([^""']+)[""']", RegexOptions.IgnoreCase);
                bool allPng = true;
                foreach (Match match in imgSrcMatches)
                {
                    string src = match.Groups[1].Value;
                    // Only consider local file references (ignore data URIs)
                    if (!src.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
                    {
                        string extension = Path.GetExtension(src);
                        Console.WriteLine($"Found image source: {src}");
                        if (!extension.Equals(".png", StringComparison.OrdinalIgnoreCase))
                        {
                            allPng = false;
                        }
                    }
                }

                Console.WriteLine(allPng
                    ? "Verification succeeded: All chart images in the HTML are PNG files."
                    : "Verification failed: Some chart images are not PNG files.");
            }
            else
            {
                Console.WriteLine("HTML file was not created.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
