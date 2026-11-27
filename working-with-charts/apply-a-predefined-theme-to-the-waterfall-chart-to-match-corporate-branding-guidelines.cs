// Title: How to apply a corporate color palette to a Waterfall chart in Aspose.Cells for .NET
// Common Searches: Aspose.Cells set custom colors for individual points in a Waterfall chart C# | apply corporate branding colors to an Excel Waterfall chart using Aspose.Cells .NET | copy theme from a template workbook to a new workbook with Aspose.Cells C# | fallback to manual series point coloring when .thmx theme not supported Aspose.Cells | save Waterfall chart with custom palette to XLSX using Aspose.Cells
// Tags: waterfall chart point color customization Aspose.Cells | manual series color assignment C# | copy workbook theme Aspose.Cells .NET | apply corporate palette to Excel chart Aspose.Cells | save workbook as xlsx Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ApplyWaterfallTheme
{
    // The example creates a workbook, adds sample data for a Waterfall chart, inserts the chart, checks for a corporate .thmx theme (which Aspose.Cells cannot apply directly), and then manually assigns specific RGB colors to each chart point (start, increase, decrease, total) to match corporate branding before saving the file as WaterfallWithCorporateTheme.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a new workbook and obtain the first worksheet.
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -----------------------------------------------------------------
                // 2. Populate sample data for a Waterfall chart.
                //    Column A – Categories, Column B – Values.
                // -----------------------------------------------------------------
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");

                sheet.Cells["A2"].PutValue("Start");
                sheet.Cells["B2"].PutValue(5000);

                sheet.Cells["A3"].PutValue("Revenue");
                sheet.Cells["B3"].PutValue(8000);

                sheet.Cells["A4"].PutValue("Cost");
                sheet.Cells["B4"].PutValue(-3000);

                sheet.Cells["A5"].PutValue("Profit");
                sheet.Cells["B5"].PutValue(5000);

                // -----------------------------------------------------------------
                // 3. Add a Waterfall chart to the worksheet.
                // -----------------------------------------------------------------
                // Parameters: chart type, upper left row, upper left column,
                // lower right row, lower right column.
                int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series (values) and categories.
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Give the chart a title.
                chart.Title.Text = "Company Waterfall – FY2025";

                // -----------------------------------------------------------------
                // 4. Attempt to apply a predefined corporate theme.
                //    Aspose.Cells does not expose a direct ApplyTheme method.
                //    If a template workbook with the desired theme is available,
                //    it can be copied via CopyTheme. Here we fall back to manual colors.
                // -----------------------------------------------------------------
                string themePath = "CorporateBranding.thmx";

                if (File.Exists(themePath))
                {
                    // Placeholder: direct theme application from .thmx is not supported.
                    // In a real scenario, load a template workbook that already uses the theme
                    // and copy its theme:
                    // Workbook template = new Workbook("TemplateWithTheme.xlsx");
                    // workbook.CopyTheme(template);
                    Console.WriteLine($"Theme file '{themePath}' found, but direct application is not supported. Proceeding with manual colors.");
                }
                else
                {
                    Console.WriteLine($"Theme file '{themePath}' not found. Applying manual colors instead.");
                }

                // -----------------------------------------------------------------
                // 5. Manual fallback – set series colors to match corporate palette.
                //    For a Waterfall chart each point can have its own color.
                // -----------------------------------------------------------------
                Series series = chart.NSeries[0];
                ChartPointCollection points = series.Points;

                // Define corporate colors (example values).
                Color startColor = Color.FromArgb(79, 129, 189);      // Accent 1
                Color increaseColor = Color.FromArgb(192, 80, 77);    // Accent 2
                Color decreaseColor = Color.FromArgb(155, 187, 89);   // Accent 3
                Color totalColor = Color.FromArgb(128, 100, 162);    // Accent 4

                // Apply colors to each point based on its role.
                // Point 0 – Start, Point 1 – Increase, Point 2 – Decrease, Point 3 – Total.
                if (points.Count >= 4)
                {
                    points[0].Area.ForegroundColor = startColor;
                    points[1].Area.ForegroundColor = increaseColor;
                    points[2].Area.ForegroundColor = decreaseColor;
                    points[3].Area.ForegroundColor = totalColor;
                }

                // -----------------------------------------------------------------
                // 6. Save the workbook.
                // -----------------------------------------------------------------
                string outputPath = "WaterfallWithCorporateTheme.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Any unexpected error is reported to the console.
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
