// Title: Aspose.Cells C# – Export Workbook to HTML with PNG Charts at 300 DPI
// Description: Demonstrates how to create a workbook, add a column chart, set HtmlSaveOptions.ExportChartImageFormat to PNG, configure the image resolution to 300 DPI, and save the file as an HTML page with crisp chart graphics.
// Keywords: Aspose.Cells C# HTML export | ExportChartImageFormat PNG | high DPI chart images | HtmlSaveOptions ImageOptions | save Excel chart as PNG HTML | Aspose.Cells chart resolution | C# workbook to HTML with charts
// Common Searches: Aspose.Cells set ExportChartImageFormat to PNG | HTML export chart resolution 300 DPI | C# save Excel workbook as HTML with PNG charts | How to increase chart image quality in Aspose.Cells HTML output | Aspose.Cells HtmlSaveOptions image DPI settings
// Developer Intent: Configure HtmlSaveOptions to output chart images as PNG at 300 DPI and generate an HTML file from a workbook.
// Use Cases: Publish financial dashboards on a web portal where column charts retain sharpness in PNG format. | Create web‑ready documentation from Excel reports, ensuring all embedded charts are high‑resolution PNGs for cross‑browser compatibility. | Automate conversion of Excel‑based analytics to HTML pages for intranet sites, preserving visual fidelity of charts.
// AI Prompts: Show C# code that sets HtmlSaveOptions.ExportChartImageFormat to PNG, defines 300 DPI image resolution, and saves a workbook with charts as HTML using Aspose.Cells. | Explain how to export an Aspose.Cells workbook to HTML with high‑quality PNG chart images, including the necessary property settings. | Provide a step‑by‑step example: create a worksheet, add a column chart, configure PNG output at 300 DPI, and write the result to an HTML file.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsHtmlExport
{
    // Demonstrates how to create a workbook, add a column chart, set HtmlSaveOptions.ExportChartImageFormat to PNG, configure the image resolution to 300 DPI, and save the file as an HTML page with crisp chart graphics.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["A3"].PutValue("Orange");
                sheet.Cells["A4"].PutValue("Banana");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(80);
                sheet.Cells["B4"].PutValue(150);

                // Add a column chart
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIdx];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Configure HTML save options with high‑resolution chart images
                HtmlSaveOptions saveOptions = new HtmlSaveOptions();
                ImageOrPrintOptions imgOpts = saveOptions.ImageOptions;
                // Default image format is PNG; only resolution is set explicitly
                imgOpts.HorizontalResolution = 300;
                imgOpts.VerticalResolution = 300;

                // Save the workbook as HTML
                string outputPath = "WorkbookWithHighResCharts.html";
                workbook.Save(outputPath, saveOptions);

                Console.WriteLine($"HTML file saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
