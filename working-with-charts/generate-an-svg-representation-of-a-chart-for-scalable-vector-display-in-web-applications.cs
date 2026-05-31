using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsSvgDemo
{
    public class GenerateChartSvg
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["B2"].PutValue(12000);
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["B3"].PutValue(15000);
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["B4"].PutValue(18000);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";      // Categories
                chart.Title.Text = "Quarterly Sales";

                // Configure SVG rendering options
                SvgImageOptions svgOptions = new SvgImageOptions
                {
                    FitToViewPort = true,               // Fit SVG to viewport for responsive scaling
                    CssPrefix = "chart-",                // Optional CSS prefix
                    EmbeddedFontType = SvgEmbeddedFontType.Woff,
                    HorizontalResolution = 300,
                    VerticalResolution = 300
                };

                // Render the chart to an SVG file
                string svgPath = "QuarterlySales.svg";
                chart.ToImage(svgPath, svgOptions);

                // Optionally save the workbook for reference
                string workbookPath = "QuarterlySales.xlsx";
                workbook.Save(workbookPath);

                Console.WriteLine($"SVG chart generated at: {Path.GetFullPath(svgPath)}");
                Console.WriteLine($"Workbook saved at: {Path.GetFullPath(workbookPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}