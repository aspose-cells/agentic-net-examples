// Title: Export an Aspose.Cells chart to JPEG with 75 % quality in C#
// Description: Creates a workbook, adds sample data and a column chart, configures ImageOrPrintOptions for JPEG at 75 % compression, and saves the chart as a JPEG file (the workbook can also be saved).
// Keywords: Aspose.Cells | C# chart export | JPEG compression | ImageOrPrintOptions | Chart.ToImage | set JPEG quality | export chart as image | .NET Excel chart image | chart to JPEG | compression quality
// Common Searches: Aspose.Cells export chart JPEG C# | set JPEG quality Aspose.Cells chart | Chart.ToImage quality parameter | ImageOrPrintOptions JPEG compression level | C# save Excel chart as JPEG with specific quality
// Developer Intent: Generate a JPEG image of a worksheet chart using Aspose.Cells with 75 % compression quality.
// Use Cases: Insert high‑quality chart images into reports while limiting file size. | Batch‑export multiple charts from a workbook with a uniform compression setting. | Create lightweight thumbnails of charts for dashboards or web previews.
// AI Prompts: Provide C# code that uses Aspose.Cells to export a chart to a JPEG file with 75 % quality. | Explain how ImageOrPrintOptions.Quality influences the size and appearance of a chart image in Aspose.Cells. | Show how to iterate over all charts in a workbook and save each as a JPEG with a custom compression level.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// Creates a workbook, adds sample data and a column chart, configures ImageOrPrintOptions for JPEG at 75 % compression, and saves the chart as a JPEG file (the workbook can also be saved).
class ExportChartToJpeg
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["A3"].PutValue("Orange");
        worksheet.Cells["A4"].PutValue("Banana");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(15);
        worksheet.Cells["B4"].PutValue(7);

        // Add a column chart
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Configure image options: JPEG format with 75% quality
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = ImageType.Jpeg;   // Specify JPEG format
        options.Quality = 75;                // Set compression quality to 75%

        // Export the chart to a JPEG image using the configured options
        chart.ToImage("ChartOutput.jpg", options);

        // Save the workbook (optional, demonstrates normal lifecycle usage)
        workbook.Save("ChartWorkbook.xlsx");
    }
}
