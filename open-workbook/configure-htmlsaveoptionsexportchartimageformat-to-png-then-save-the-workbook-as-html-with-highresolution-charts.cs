using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;   // For ImageType enum

namespace AsposeCellsHtmlExport
{
    class Program
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

                // Configure HTML save options
                HtmlSaveOptions saveOptions = new HtmlSaveOptions();

                // Access the ImageOrPrintOptions through HtmlSaveOptions.ImageOptions
                ImageOrPrintOptions imgOpts = saveOptions.ImageOptions;

                // Set chart image format to PNG (default, but set explicitly)
                imgOpts.ImageType = ImageType.Png;

                // Set high resolution for chart images
                imgOpts.HorizontalResolution = 300;
                imgOpts.VerticalResolution = 300;

                // Embed images as Base64 to keep a single HTML file
                saveOptions.ExportImagesAsBase64 = true;

                // Save the workbook as HTML with the configured options
                workbook.Save("HighResChart.html", saveOptions);

                Console.WriteLine("Workbook saved as HTML with high‑resolution PNG charts.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}