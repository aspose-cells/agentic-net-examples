using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsSvgExportDemo
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
                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(210);
                sheet.Cells["B4"].PutValue(150);

                // Add a line chart covering the data range
                int chartIdx = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIdx];
                chart.NSeries.Add("B2:B4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";      // Categories
                chart.Title.Text = "Quarterly Sales";

                // Configure SVG rendering options (no ImageFormat needed for SVG)
                SvgImageOptions svgOpts = new SvgImageOptions
                {
                    FitToViewPort = true,            // Make SVG fit the viewport
                    CssPrefix = "chart-",            // Optional: custom CSS prefix
                    EmbeddedFontType = SvgEmbeddedFontType.Woff // Optional: embed fonts
                };

                // Export the chart directly to an SVG file
                string svgPath = "QuarterlySalesChart.svg";

                // Ensure the directory for the SVG exists
                string svgDir = Path.GetDirectoryName(svgPath);
                if (!string.IsNullOrEmpty(svgDir) && !Directory.Exists(svgDir))
                {
                    Directory.CreateDirectory(svgDir);
                }

                try
                {
                    chart.ToImage(svgPath, svgOpts);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to export chart to SVG: {ex.Message}");
                }

                // Optionally, save the workbook for reference
                string workbookPath = "QuarterlySalesWorkbook.xlsx";
                try
                {
                    workbook.Save(workbookPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }

                Console.WriteLine($"Chart exported to SVG: {Path.GetFullPath(svgPath)}");
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(workbookPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}