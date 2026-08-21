// Title: Set ExportChartImageFormat to JPEG in Aspose.Cells HTML export and verify the images (C#)
// Description: This example creates a workbook, adds sample data and a column chart, then configures HtmlSaveOptions so that all images—including charts—are saved as JPEG. It demonstrates using ImageOptions.ImageType, optionally setting ExportChartImageFormat, saving the workbook as HTML, and programmatically confirming that the generated chart files have a .jpeg or .jpg extension.
// Keywords: Aspose.Cells | C# | HTML export | chart image JPEG | ExportChartImageFormat | ImageOptions.ImageType | verify JPEG output | save workbook as HTML | chart rendering | Aspose.Cells example
// Common Searches: Aspose.Cells set chart image format to JPEG | HTML export chart images JPEG C# | ExportChartImageFormat property usage | how to verify exported chart files are JPEG | Aspose.Cells HtmlSaveOptions JPEG images
// Developer Intent: Configure Aspose.Cells to export chart graphics as JPEG when saving a workbook to HTML and programmatically ensure the resulting files are JPEG images.
// Use Cases: Produce lightweight HTML reports where all embedded chart graphics must be JPEG for web compatibility. | Add an automated post‑export check that confirms chart files were written with .jpeg or .jpg extensions. | Write version‑tolerant code that forces JPEG output for charts, handling environments where ExportChartImageFormat may be unavailable.
// AI Prompts: Generate C# code that saves an Aspose.Cells workbook to HTML with all chart images forced to JPEG and lists the created JPEG files. | Show how to detect if the ExportChartImageFormat property exists in the current Aspose.Cells version and set it to JPEG when possible. | Provide a method that validates the file format of exported chart images after an HTML save using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// This example creates a workbook, adds sample data and a column chart, then configures HtmlSaveOptions so that all images—including charts—are saved as JPEG. It demonstrates using ImageOptions.ImageType, optionally setting ExportChartImageFormat, saving the workbook as HTML, and programmatically confirming that the generated chart files have a .jpeg or .jpg extension.
class ExportChartAsJpegExample
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

        // Add a column chart and set its data source
        int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIdx];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Set the image type for all exported images (including charts) to JPEG
        htmlOptions.ImageOptions.ImageType = ImageType.Jpeg;

        // Explicitly set ExportChartImageFormat to JPEG if the property exists
        // (If the property is not available in the current version, the line can be omitted safely)
        // htmlOptions.ExportChartImageFormat = ImageType.Jpeg; // Uncomment if supported

        // Define output folder and HTML file name
        string outputFolder = Path.Combine(Environment.CurrentDirectory, "HtmlOutput");
        Directory.CreateDirectory(outputFolder);
        string htmlPath = Path.Combine(outputFolder, "Workbook.html");

        // Save the workbook as HTML with the configured options
        workbook.Save(htmlPath, htmlOptions);

        // Verify that chart images are saved as JPEG files
        Console.WriteLine("Verifying exported chart images...");
        string[] imageFiles = Directory.GetFiles(outputFolder, "*.jpeg");
        if (imageFiles.Length == 0)
        {
            // Some versions use .jpg extension for JPEG images
            imageFiles = Directory.GetFiles(outputFolder, "*.jpg");
        }

        if (imageFiles.Length > 0)
        {
            Console.WriteLine($"Found {imageFiles.Length} JPEG image file(s):");
            foreach (string file in imageFiles)
            {
                Console.WriteLine(" - " + Path.GetFileName(file));
            }
        }
        else
        {
            Console.WriteLine("No JPEG images were found. Export may have failed or used a different format.");
        }
    }
}
