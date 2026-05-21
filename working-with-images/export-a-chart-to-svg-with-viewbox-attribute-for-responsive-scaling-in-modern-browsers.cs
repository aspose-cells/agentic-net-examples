using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class ExportChartToResponsiveSvg
{
    static void Main()
    {
        // ---------- Create a new workbook ----------
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // ---------- Populate worksheet with sample data ----------
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["A3"].PutValue("Orange");
        worksheet.Cells["A4"].PutValue("Banana");

        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(120);
        worksheet.Cells["B3"].PutValue(80);
        worksheet.Cells["B4"].PutValue(150);

        // ---------- Add a column chart ----------
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories
        chart.Title.Text = "Fruit Sales";

        // ---------- Configure SVG rendering options ----------
        SvgImageOptions svgOptions = new SvgImageOptions();
        svgOptions.ImageType = ImageType.Svg;      // Ensure output format is SVG
        svgOptions.FitToViewPort = true;          // Adds viewBox attribute for responsive scaling

        // ---------- Export the chart to an SVG file ----------
        // The file will contain a viewBox attribute, allowing the SVG to scale with the viewport.
        chart.ToImage("FruitSales.svg", svgOptions);

        Console.WriteLine("Chart exported to SVG with responsive viewBox.");
    }
}