// Title: Export Aspose.Cells Chart to Scalable SVG in C#
// Description: Shows how to build a workbook, add a column chart, set SvgImageOptions (FitToViewPort, CSS prefix, embedded WOFF font) and save the chart as an SVG vector file with Aspose.Cells for .NET, preserving resolution‑independent rendering for web pages.
// Keywords: Aspose.Cells | C# chart export | SVG vector | chart.ToImage | SvgImageOptions | Excel chart SVG | scalable web graphics | vector chart .NET | Aspose.Cells SVG example | embed chart SVG
// Common Searches: Aspose.Cells export chart to SVG C# | How to save Excel chart as SVG using Aspose.Cells | SvgImageOptions example Aspose.Cells | C# generate SVG chart from workbook | Export Aspose.Cells chart as vector graphic
// Developer Intent: Generate an SVG file from an Aspose.Cells chart so the graphic stays fully scalable when embedded in web pages or applications.
// Use Cases: Create responsive dashboards by converting Excel charts to SVG for HTML5 rendering. | Automate batch export of all workbook charts to individual SVG files with custom CSS prefixes. | Produce printable reports where charts retain crisp quality at any zoom level. | Integrate SVG charts into mobile apps without pixelation.
// AI Prompts: Provide C# code that iterates through every chart in an Aspose.Cells workbook and saves each as an SVG file with unique filenames. | Show how to export a pie chart to SVG with a transparent background and embedded WOFF fonts using Aspose.Cells. | Explain how to adjust SvgImageOptions to control viewBox and CSS styling for exported charts. | Generate a step‑by‑step guide for converting a line chart to SVG and embedding it in an ASP.NET page.

using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using System;

// Shows how to build a workbook, add a column chart, set SvgImageOptions (FitToViewPort, CSS prefix, embedded WOFF font) and save the chart as an SVG vector file with Aspose.Cells for .NET, preserving resolution‑independent rendering for web pages.
class ExportChartToSvg
{
    public static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["A3"].PutValue("Orange");
            worksheet.Cells["A4"].PutValue("Banana");
            worksheet.Cells["B1"].PutValue("Sales");
            worksheet.Cells["B2"].PutValue(120);
            worksheet.Cells["B3"].PutValue(80);
            worksheet.Cells["B4"].PutValue(150);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";      // Categories
            chart.Title.Text = "Sales by Fruit";

            // Configure SVG rendering options
            SvgImageOptions svgOptions = new SvgImageOptions
            {
                FitToViewPort = true,               // Make SVG fit the viewport
                CssPrefix = "chart-",                // Optional CSS prefix
                EmbeddedFontType = SvgEmbeddedFontType.Woff // Embed WOFF font if needed
            };

            // Export the chart as an SVG file (vector graphic)
            chart.ToImage("sales_chart.svg", svgOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
