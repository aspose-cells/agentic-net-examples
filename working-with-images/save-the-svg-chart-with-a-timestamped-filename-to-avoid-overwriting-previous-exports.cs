// Title: Export Aspose.Cells Chart to SVG with a Timestamped Filename (C#)
// Description: Creates a workbook, adds sample data and a column chart, configures SvgImageOptions, generates a filename that includes the current date‑time (yyyyMMdd_HHmmss), and saves the chart as an SVG file using Chart.ToImage. Each export receives a unique name, preventing overwrites.
// Keywords: Aspose.Cells | C# | chart export | SVG | timestamped filename | unique file name | Chart.ToImage | SvgImageOptions | avoid overwriting | .NET
// Common Searches: Aspose.Cells export chart to SVG C# | save chart as SVG with timestamp | unique filename for Aspose.Cells chart export | Chart.ToImage avoid overwriting files | generate date‑time suffix for SVG file Aspose.Cells
// Developer Intent: Generate a uniquely named SVG file for a chart to prevent overwriting previous exports.
// Use Cases: Automated daily sales dashboards that archive each SVG chart with a date‑time suffix. | Batch processing of multiple workbook charts where each SVG file must have a distinct name. | CI/CD pipelines that export charts during build runs and need version‑controlled filenames.
// AI Prompts: Write a reusable C# method that accepts a Chart object and saves it as an SVG file with a customizable timestamp format. | Show how to modify the filename pattern to include milliseconds for high‑frequency chart exports using Aspose.Cells. | Explain how to integrate timestamped SVG chart export into an ASP.NET Core reporting service.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// Creates a workbook, adds sample data and a column chart, configures SvgImageOptions, generates a filename that includes the current date‑time (yyyyMMdd_HHmmss), and saves the chart as an SVG file using Chart.ToImage. Each export receives a unique name, preventing overwrites.
class ExportChartSvg
{
    static void Main()
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
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(15);
        worksheet.Cells["B4"].PutValue(7);

        // Add a column chart
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Prepare SVG rendering options
        SvgImageOptions svgOptions = new SvgImageOptions
        {
            ImageType = ImageType.Svg,   // Ensure SVG output
            FitToViewPort = true        // Optional: fit SVG to viewport
        };

        // Generate a timestamped filename to avoid overwriting previous exports
        string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        string svgFileName = $"ChartExport_{timestamp}.svg";

        // Export the chart to SVG using the timestamped filename
        chart.ToImage(svgFileName, svgOptions);

        Console.WriteLine($"Chart exported successfully to: {svgFileName}");
    }
}
