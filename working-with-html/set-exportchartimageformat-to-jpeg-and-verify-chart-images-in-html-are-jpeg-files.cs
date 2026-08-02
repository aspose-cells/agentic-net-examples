// Title: Export Excel Chart as JPEG in HTML using Aspose.Cells for .NET and Verify Files
// Description: Demonstrates how to create a workbook, add a column chart, configure HtmlSaveOptions to save chart images as separate JPEG files, export the workbook to HTML, and programmatically confirm that the generated image files have a .jpg or .jpeg extension.
// Keywords: Aspose.Cells | C# | .NET | HTML export | chart image JPEG | ImageOptions.ImageType | ExportChartImageFormat | verify chart image format | separate image files | Aspose.Cells HTMLSaveOptions
// Common Searches: Aspose.Cells export chart to JPEG HTML | How to set chart image format to JPEG in Aspose.Cells | Verify chart images are JPEG after HTML export .NET | Save Excel chart as JPEG file using Aspose.Cells | HtmlSaveOptions ImageType JPEG example
// Developer Intent: Configure Aspose.Cells to output chart images as JPEG files during HTML conversion and programmatically ensure the saved images use the JPEG extension.
// Use Cases: Create web‑ready HTML reports where chart graphics are stored as JPEG files for reduced size and broad browser support. | Automate a CI/CD check that confirms exported chart images meet a JPEG format requirement. | Generate separate JPEG chart images for embedding in newsletters, blogs, or external web pages.
// AI Prompts: Write C# code with Aspose.Cells that exports a workbook to HTML, saves chart images as JPEG files, and validates the file extensions. | Explain the effect of HtmlSaveOptions.ImageOptions.ImageType on chart rendering and show how to detect JPEG files after export. | Modify the example to output chart images as PNG while keeping them as external files.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// Demonstrates how to create a workbook, add a column chart, configure HtmlSaveOptions to save chart images as separate JPEG files, export the workbook to HTML, and programmatically confirm that the generated image files have a .jpg or .jpeg extension.
class ExportChartAsJpegDemo
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

        // Add a column chart and bind the data
        int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIdx];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        // Ensure images are saved as separate files (not Base64)
        htmlOptions.ExportImagesAsBase64 = false;
        // Set the image type for charts (and other images) to JPEG
        htmlOptions.ImageOptions.ImageType = ImageType.Jpeg;

        // Define output paths
        string outputFolder = Path.Combine(Environment.CurrentDirectory, "HtmlOutput");
        Directory.CreateDirectory(outputFolder);
        string htmlPath = Path.Combine(outputFolder, "Workbook.html");

        // Save the workbook as HTML
        workbook.Save(htmlPath, htmlOptions);
        Console.WriteLine($"HTML saved to: {htmlPath}");

        // The HTML save creates a subfolder with images (Workbook_files)
        string imagesFolder = Path.Combine(outputFolder, "Workbook_files");
        if (Directory.Exists(imagesFolder))
        {
            // Get all image files generated for the chart
            string[] imageFiles = Directory.GetFiles(imagesFolder);
            Console.WriteLine($"Found {imageFiles.Length} image file(s) in '{imagesFolder}':");
            foreach (string imgPath in imageFiles)
            {
                Console.WriteLine($" - {Path.GetFileName(imgPath)} (Extension: {Path.GetExtension(imgPath)})");
            }

            // Verify that at least one image has a JPEG extension
            bool hasJpeg = false;
            foreach (string imgPath in imageFiles)
            {
                string ext = Path.GetExtension(imgPath).ToLowerInvariant();
                if (ext == ".jpg" || ext == ".jpeg")
                {
                    hasJpeg = true;
                    break;
                }
            }

            Console.WriteLine(hasJpeg
                ? "Verification passed: Chart images are saved as JPEG."
                : "Verification failed: No JPEG images were found.");
        }
        else
        {
            Console.WriteLine("No image folder was created. Verification cannot be performed.");
        }
    }
}
