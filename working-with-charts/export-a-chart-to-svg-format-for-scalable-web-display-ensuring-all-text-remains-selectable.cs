using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class ExportChartToSvg
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
        worksheet.Cells["B2"].PutValue(120);
        worksheet.Cells["B3"].PutValue(80);
        worksheet.Cells["B4"].PutValue(150);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);               // Values
        chart.NSeries.CategoryData = "A2:A4";           // Categories
        chart.Title.Text = "Fruit Sales";

        // Configure SVG rendering options
        SvgImageOptions svgOptions = new SvgImageOptions();
        svgOptions.ImageType = ImageType.Svg;           // Output format: SVG
        svgOptions.FitToViewPort = true;               // Fit SVG to viewport for responsive web display
        svgOptions.CssPrefix = "chart-";                // Optional CSS prefix
        svgOptions.EmbeddedFontType = SvgEmbeddedFontType.Woff; // Embed font so text remains selectable

        // Export the chart as an SVG file
        chart.ToImage("FruitSales.svg", svgOptions);
    }
}