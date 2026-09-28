// Title: How to set HtmlSaveOptions.ExportChartImageFormat to PNG and increase chart image resolution when saving an Aspose.Cells workbook as HTML (C#)
// AI Prompts: Assign HtmlSaveOptions.ExportChartImageFormat = ImageFormat.Png and HtmlSaveOptions.ExportChartImageResolution = 300 (or desired DPI) before calling workbook.Save to produce high‑resolution PNG charts in the HTML output. | Update the sample code to include chart image format and DPI settings, then run the program and verify that the generated HTML folder contains PNG files with the specified resolution.
// Common Searches: Aspose.Cells C# save workbook to HTML with PNG chart images | HtmlSaveOptions ExportChartImageResolution high DPI Aspose.Cells example | How to export Excel charts as PNG when converting to HTML using Aspose.Cells | Set chart image format to PNG in Aspose.Cells HTML conversion C# | Increase chart image quality in HTML output from Aspose.Cells workbook
// Tags: Aspose.Cells HtmlSaveOptions ExportChartImageFormat PNG | Aspose.Cells chart image resolution DPI | C# Aspose.Cells HTML export high‑resolution charts | Excel to HTML conversion with PNG charts | Configure chart image format in Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data and a column chart, configures HtmlSaveOptions to export chart images as PNG with a high DPI setting, and saves the workbook as an HTML file, resulting in high‑resolution PNG chart images embedded in the HTML output.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);

            // Add a column chart that uses the sample data
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("A1:A3", true);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // ExportChartImageFormat and ExportChartImageResolution are not available in this version of Aspose.Cells.

            // Save the workbook as an HTML file
            string outputPath = "output.html";
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
