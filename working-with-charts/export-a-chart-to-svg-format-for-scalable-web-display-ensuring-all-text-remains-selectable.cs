// Title: Export Aspose.Cells Chart to SVG with Selectable Text (C#)
// Description: Creates a workbook, adds sample fruit‑sales data, builds a column chart, and uses SvgImageOptions (embedded WOFF font, viewport fit, optional CSS prefix) to export the chart as an SVG file where all text remains selectable. The workbook can also be saved as an XLSX file.
// Keywords: Aspose.Cells SVG export | C# chart to SVG | selectable text SVG | SvgImageOptions embed font | responsive SVG chart | Excel chart SVG .NET | chart.ToImage SVG
// Common Searches: Aspose.Cells export chart to SVG C# | how to keep text selectable in SVG chart | embed fonts in SVG using Aspose.Cells | fit SVG chart to viewport Aspose | export Excel column chart as SVG
// Developer Intent: Generate an SVG version of an Aspose.Cells chart while preserving selectable, searchable text.
// Use Cases: Render scalable, searchable charts for web dashboards directly from Excel data. | Produce print‑ready SVG graphics with embedded fonts for high‑quality reports. | Deliver responsive chart images for mobile sites without sacrificing accessibility.
// AI Prompts: Write C# code with Aspose.Cells to export a pie chart to SVG and embed fonts so the labels stay selectable. | Explain how to configure SvgImageOptions for viewport fitting and custom CSS prefixes when exporting charts. | Show how to iterate through all charts in a workbook and save each as a separate SVG file using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample fruit‑sales data, builds a column chart, and uses SvgImageOptions (embedded WOFF font, viewport fit, optional CSS prefix) to export the chart as an SVG file where all text remains selectable. The workbook can also be saved as an XLSX file.
    public class ExportChartToSvg
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
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
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";     // Categories
            chart.Title.Text = "Fruit Sales";

            // Configure SVG rendering options (ImageFormat is implicit for SvgImageOptions)
            SvgImageOptions svgOptions = new SvgImageOptions
            {
                FitToViewPort = true,                   // Fit SVG to viewport for responsive display
                CssPrefix = "chart-",                   // Optional CSS prefix
                EmbeddedFontType = SvgEmbeddedFontType.Woff // Embed font so text remains selectable
            };

            // Export the chart directly to an SVG file
            string svgPath = "FruitSalesChart.svg";
            try
            {
                chart.ToImage(svgPath, svgOptions);
                Console.WriteLine($"Chart exported to {Path.GetFullPath(svgPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to export chart: {ex.Message}");
            }

            // (Optional) Save the workbook if you need the original Excel file
            string workbookPath = "FruitSalesWorkbook.xlsx";
            try
            {
                workbook.Save(workbookPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(workbookPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
