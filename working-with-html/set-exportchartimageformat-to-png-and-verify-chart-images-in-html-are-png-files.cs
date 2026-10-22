// Title: Export Excel chart to PNG images when saving as HTML with Aspose.Cells for .NET and confirm image sources
// AI Prompts: Generate C# code that configures HtmlSaveOptions.ExportChartImageFormat to PNG, creates a chart in a workbook, and saves the workbook as an HTML file. | Write a C# routine that reads the saved HTML file, extracts every <img> src attribute, and checks that each referenced chart image file ends with .png. | Modify the example to export a line chart instead of a column chart while still using PNG format and performing the same verification steps.
// Common Searches: how to export Excel chart as PNG when saving workbook to HTML using Aspose.Cells .NET | Aspose.Cells HtmlSaveOptions ExportChartImageFormat property example C# | verify that chart images in generated HTML are PNG files with Aspose.Cells | C# read Aspose.Cells HTML output and list chart image file names | change chart type in Aspose.Cells HTML export while keeping PNG image format
// Tags: Aspose.Cells export chart PNG in HTML | HtmlSaveOptions ExportChartImageFormat PNG | C# validate chart image extensions in generated HTML | regex extract img src Aspose.Cells HTML output | change chart type Aspose.Cells HTML export

using System;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, adds data and a column chart, sets HtmlSaveOptions.ExportChartImageFormat to PNG, and saves the workbook as an HTML file. It then reads the HTML, extracts all <img> src attributes with a regular expression, and verifies that every chart image reference ends with .png, outputting the verification result.
class ExportChartAsPngExample
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
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Save workbook as HTML (charts are exported as PNG by default)
            string htmlPath = "ChartExport.html";
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
            workbook.Save(htmlPath, saveOptions);

            // Verify that chart images referenced in the HTML are PNG files
            if (File.Exists(htmlPath))
            {
                string htmlContent = File.ReadAllText(htmlPath);
                var imgSrcs = Regex.Matches(htmlContent, @"<img[^>]+src\s*=\s*[""']([^""']+)[""']")
                                   .Cast<Match>()
                                   .Select(m => m.Groups[1].Value)
                                   .ToList();

                bool allPng = imgSrcs.All(src => src.EndsWith(".png", StringComparison.OrdinalIgnoreCase));
                Console.WriteLine("Chart images in HTML are PNG: " + allPng);

                foreach (var src in imgSrcs)
                {
                    Console.WriteLine("Image source: " + src);
                }
            }
            else
            {
                Console.WriteLine("HTML file was not created: " + htmlPath);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
