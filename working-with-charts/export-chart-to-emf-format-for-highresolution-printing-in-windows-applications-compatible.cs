// Title: How to export an Excel column chart to a high‑resolution EMF image using Aspose.Cells for C#
// AI Prompts: Generate C# code that creates a column chart from worksheet data and saves it as a 300 DPI EMF file with Aspose.Cells. | Show how to configure ImageOrPrintOptions to use ImageType.Emf and EmfRenderSetting.EmfOnly for chart‑only EMF export. | Adapt the example to export a line chart at 600 DPI in EMF format using Aspose.Cells.
// Common Searches: Aspose.Cells C# export chart as EMF with custom DPI | Create high resolution EMF image from Excel chart using Aspose.Cells | How to set EmfRenderSetting to EmfOnly in Aspose.Cells chart export | Export column chart to EMF file for Windows printing with Aspose.Cells
// Tags: chart export to EMF Aspose.Cells C# | ImageOrPrintOptions EMF DPI configuration | EmfRenderSetting.EmfOnly usage | high‑resolution chart image generation | column chart EMF output example

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// The sample creates a workbook, fills cells with fruit sales data, adds a column chart, configures ImageOrPrintOptions for 300 DPI EMF output using EmfRenderSetting.EmfOnly, and saves the chart as 'FruitSales.emf' for high‑resolution Windows printing.
class ExportChartToEmf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["A3"].PutValue("Banana");
        worksheet.Cells["A4"].PutValue("Cherry");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(120);
        worksheet.Cells["B3"].PutValue(80);
        worksheet.Cells["B4"].PutValue(150);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);               // Values
        chart.NSeries.CategoryData = "A2:A4";           // Categories
        chart.Title.Text = "Fruit Sales";

        // Configure image options for high‑resolution EMF output
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Emf,                  // Export as EMF
            HorizontalResolution = 300,                 // 300 DPI horizontal
            VerticalResolution = 300,                   // 300 DPI vertical
            EmfRenderSetting = EmfRenderSetting.EmfOnly // Render EMF records only
        };

        // Export the chart to an EMF file
        chart.ToImage("FruitSales.emf", options);

        Console.WriteLine("Chart exported to EMF successfully.");
    }
}
