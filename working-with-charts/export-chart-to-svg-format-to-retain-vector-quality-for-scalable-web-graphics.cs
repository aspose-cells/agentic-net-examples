// Title: Export Aspose.Cells Chart to SVG in C# – Preserve Vector Quality
// Description: Creates a workbook, adds sample data and a line chart, configures SvgImageOptions (FitToViewPort, CssPrefix), and uses chart.ToImage to generate a vector‑based SVG file. The workbook can be saved optionally for reference.
// Keywords: Aspose.Cells SVG export | C# chart to SVG | SvgImageOptions | vector chart export .NET | Aspose.Cells line chart SVG | chart.ToImage SVG
// Common Searches: Aspose.Cells export chart SVG C# | How to save chart as SVG using Aspose.Cells | SvgImageOptions example Aspose.Cells | C# generate scalable SVG chart | Export Excel chart to SVG with Aspose
// Developer Intent: Generate an SVG file from an Aspose.Cells chart while keeping the output fully vectorized.
// Use Cases: Embed a high‑resolution line‑chart SVG in a responsive web dashboard. | Create SVG chart assets for email templates that require scalable graphics. | Produce vector charts for print‑ready PDFs or high‑DPI reports.
// AI Prompts: Provide C# code to export a pie chart to SVG with a custom CSS prefix using Aspose.Cells. | Show how to iterate through all charts in a workbook and save each as a separate SVG with different rendering options. | Explain how to adjust the viewport size and apply CSS styling when rendering a chart to SVG with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// Creates a workbook, adds sample data and a line chart, configures SvgImageOptions (FitToViewPort, CssPrefix), and uses chart.ToImage to generate a vector‑based SVG file. The workbook can be saved optionally for reference.
class ExportChartToSvg
{
    static void Main()
    {
        try
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
            chart.Title.Text = "Monthly Sales";

            // Configure SVG rendering options
            SvgImageOptions svgOptions = new SvgImageOptions
            {
                FitToViewPort = true,      // Fit SVG to viewport (optional)
                CssPrefix = "chart-"        // Custom CSS prefix (optional)
            };

            // Export the chart to an SVG file while retaining vector quality
            string svgPath = "output_chart.svg";
            chart.ToImage(svgPath, svgOptions);
            Console.WriteLine($"Chart exported to SVG: {Path.GetFullPath(svgPath)}");

            // Save the workbook (optional, to keep the source data)
            string workbookPath = "ChartWorkbook.xlsx";
            workbook.Save(workbookPath);
            Console.WriteLine($"Workbook saved: {Path.GetFullPath(workbookPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
