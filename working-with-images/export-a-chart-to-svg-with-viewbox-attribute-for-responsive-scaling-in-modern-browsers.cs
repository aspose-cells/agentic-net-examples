// Title: Export Excel Chart to Responsive SVG with viewBox using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add a column chart, configure SvgImageOptions with FitToViewPort (adds a viewBox attribute) and an optional CSS prefix, then export the chart to an SVG file while optionally saving the original XLSX.
// Keywords: Aspose.Cells | C# | .NET | export chart to SVG | FitToViewPort | viewBox | responsive SVG | SvgImageOptions | CSS prefix | Excel chart SVG | column chart export | web dashboard graphics
// Common Searches: Aspose.Cells export chart to SVG with viewBox | C# FitToViewPort SvgImageOptions example | How to add viewBox attribute when exporting Excel chart to SVG | Set CSS prefix for SVG chart export Aspose.Cells | Responsive SVG chart from Excel using Aspose.Cells
// Developer Intent: Generate an SVG representation of an Excel chart that includes a viewBox attribute for fluid scaling in browsers.
// Use Cases: Embed a scalable SVG chart in web dashboards without losing aspect ratio. | Export multiple charts to SVG on the same page while avoiding CSS naming conflicts. | Provide both an SVG image for front‑end display and the original XLSX for archival or further analysis.
// AI Prompts: Write C# code that exports a pie chart from an Aspose.Cells workbook to SVG with FitToViewPort enabled and a custom CSS prefix. | Explain how the FitToViewPort property inserts a viewBox attribute into the SVG and why this enables responsive scaling. | Show how to embed the exported SVG chart into an HTML page and make it resize automatically using CSS.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsSvgExport
{
    // Demonstrates how to create a workbook, add a column chart, configure SvgImageOptions with FitToViewPort (adds a viewBox attribute) and an optional CSS prefix, then export the chart to an SVG file while optionally saving the original XLSX.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
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

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";     // Categories
                chart.Title.Text = "Quarterly Sales";

                // Configure SVG rendering options
                SvgImageOptions svgOptions = new SvgImageOptions
                {
                    // Ensure the generated SVG fits the viewport (adds viewBox attribute)
                    FitToViewPort = true,

                    // Optional: add a CSS prefix to avoid naming collisions
                    CssPrefix = "chart-"
                };

                // Export the chart to an SVG file using the configured options
                chart.ToImage("QuarterlySales.svg", svgOptions);

                // Save the workbook if you also want the Excel file
                workbook.Save("QuarterlySales.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
