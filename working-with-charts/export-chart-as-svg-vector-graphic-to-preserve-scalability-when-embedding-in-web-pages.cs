using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

class ExportChartToSvg
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

            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["B3"].PutValue(800);
            sheet.Cells["B4"].PutValue(1500);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";      // Categories
            chart.Title.Text = "Sales Chart";

            // Configure SVG rendering options
            SvgImageOptions svgOptions = new SvgImageOptions
            {
                // Make SVG fit the viewport
                FitToViewPort = true,
                // Optional CSS prefix for styling
                CssPrefix = "chart-"
            };

            // Export the chart as an SVG file (vector graphic)
            string svgPath = "sales_chart.svg";
            chart.ToImage(svgPath, svgOptions);

            // Save the workbook (optional, for reference)
            string workbookPath = "SalesWorkbook.xlsx";
            workbook.Save(workbookPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}