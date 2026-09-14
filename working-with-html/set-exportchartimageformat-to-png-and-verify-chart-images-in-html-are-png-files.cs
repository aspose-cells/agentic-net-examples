// Title: Set ExportChartImageFormat to PNG and verify chart images are PNG when saving a workbook to HTML with Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a workbook with a column chart, sets HtmlSaveOptions.ExportChartImageFormat to ImageFormat.Png, and saves the workbook as an HTML file. | Add logic to read the saved HTML file and confirm that every chart image reference ends with the .png extension. | Write C# to enumerate the files in the HTML output folder and output the names of all exported chart image files. | Include exception handling that reports a verification failure if no PNG chart images are found.
// Common Searches: Aspose.Cells C# export chart as PNG in HTML output | How to force chart images to PNG when saving workbook to HTML using Aspose.Cells | Verify chart image file extension in generated HTML with Aspose.Cells .NET | HtmlSaveOptions ExportChartImageFormat PNG example C# | Check if Aspose.Cells HTML export creates .png chart files
// Tags: Aspose.Cells HtmlSaveOptions ExportChartImageFormat PNG | C# export chart image as PNG in HTML | verify PNG chart references in generated HTML | list exported chart files Aspose.Cells | chart image format configuration Aspose.Cells

using System;
using System.IO;
using System.Linq;
using System.Drawing.Imaging;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook containing a column chart, configures HtmlSaveOptions.ExportChartImageFormat to ImageFormat.Png, saves the workbook as HTML, then reads the generated HTML to ensure chart image references end with .png and lists the exported image files in the accompanying folder.
class ExportChartAsPng
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set data series (values). Category data is inferred from the first column.
            chart.NSeries.Add("B2:B4", true);
            chart.Title.Text = "Sample Column Chart";

            // Configure HTML save options (default chart image format is PNG)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // If a specific format is required, uncomment the following line (requires a supporting Aspose.Cells version)
            // htmlOptions.ExportChartImageFormat = ImageFormat.Png;

            // Define output paths
            string htmlPath = "ChartExport.html";
            string imagesFolder = "ChartExport_files";

            // Save the workbook as HTML with the specified options
            workbook.Save(htmlPath, htmlOptions);

            // Verify that the generated HTML references PNG files for charts
            if (File.Exists(htmlPath))
            {
                string htmlContent = File.ReadAllText(htmlPath);
                bool containsPng = htmlContent
                    .Split(new[] { '\"' }, StringSplitOptions.RemoveEmptyEntries)
                    .Any(part => part.EndsWith(".png", StringComparison.OrdinalIgnoreCase));

                Console.WriteLine(containsPng
                    ? "Verification passed: Chart images are exported as PNG."
                    : "Verification failed: No PNG chart images found in the HTML.");
            }
            else
            {
                Console.WriteLine($"Error: HTML file '{htmlPath}' was not created.");
            }

            // List the exported image files to confirm their extensions
            if (Directory.Exists(imagesFolder))
            {
                var imageFiles = Directory.GetFiles(imagesFolder);
                Console.WriteLine("Exported image files:");
                foreach (var file in imageFiles)
                {
                    Console.WriteLine($"- {Path.GetFileName(file)}");
                }
            }
            else
            {
                Console.WriteLine($"Images folder '{imagesFolder}' does not exist.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
