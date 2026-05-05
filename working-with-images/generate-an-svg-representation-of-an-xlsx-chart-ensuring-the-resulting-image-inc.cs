using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using System;

class GenerateChartSvg
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Month");
        worksheet.Cells["A2"].PutValue("Jan");
        worksheet.Cells["A3"].PutValue("Feb");
        worksheet.Cells["A4"].PutValue("Mar");

        worksheet.Cells["B1"].PutValue("Sales");
        worksheet.Cells["B2"].PutValue(120);
        worksheet.Cells["B3"].PutValue(210);
        worksheet.Cells["B4"].PutValue(150);

        // Add a line chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Configure SVG rendering options
        SvgImageOptions svgOptions = new SvgImageOptions
        {
            FitToViewPort = true          // Generates SVG with a proper viewBox attribute
        };

        // Render the chart to an SVG file
        chart.ToImage("chart_output.svg", svgOptions);
    }
}