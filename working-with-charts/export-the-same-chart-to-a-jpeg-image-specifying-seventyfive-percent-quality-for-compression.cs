// Title: Export Aspose.Cells Chart to JPEG with 75% Quality (C#)
// Description: Creates a workbook, adds sample data, builds a column chart, and uses ImageOrPrintOptions to export the chart as a JPEG file with a 75 % compression quality. The example also shows optional workbook saving.
// Keywords: Aspose.Cells chart export JPEG | C# Aspose.Cells ImageOrPrintOptions | JPEG quality 75 Aspose.Cells | chart.ToImage compression | .NET chart to JPEG example | Aspose.Cells image options | export chart as JPEG C# | Aspose.Cells chart image quality
// Common Searches: Aspose.Cells export chart JPEG | C# set JPEG quality Aspose.Cells | ImageOrPrintOptions quality 75 | chart.ToImage JPEG compression | how to save chart as JPEG in Aspose.Cells
// Developer Intent: Save a chart from an Aspose.Cells workbook as a JPEG file while specifying a 75 % compression quality.
// Use Cases: Generate lightweight chart thumbnails for web pages with controlled file size. | Create printable chart images where quality and document size must be balanced. | Automate extraction of charts from multiple workbooks and store them as JPEGs with consistent compression.
// AI Prompts: Show C# code to export an Aspose.Cells chart to PNG with a custom DPI. | Explain how to adjust the JPEG quality value when using ImageOrPrintOptions in Aspose.Cells. | Provide a script that batch‑exports all charts in a workbook to JPEG files with varying quality settings.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// Creates a workbook, adds sample data, builds a column chart, and uses ImageOrPrintOptions to export the chart as a JPEG file with a 75 % compression quality. The example also shows optional workbook saving.
class ExportChartToJpeg
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
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(15);
        sheet.Cells["B4"].PutValue(7);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Configure image options: JPEG format with 75% quality
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = ImageType.Jpeg;   // Set output format to JPEG
        options.Quality = 75;                 // Compression quality (0-100)

        // Export the chart to a JPEG file using the configured options
        chart.ToImage("ChartOutput.jpg", options);

        // Save the workbook (optional, demonstrates full lifecycle)
        workbook.Save("ChartWorkbook.xlsx");
    }
}
