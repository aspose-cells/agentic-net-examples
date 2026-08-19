// Title: Export a Line Chart to SVG using Aspose.Cells for .NET (C#)
// Description: The sample creates a workbook, fills cells A1:B4 with month and sales values, adds a line chart, sets SvgImageOptions (FitToViewPort, custom CSS prefix, WOFF font embedding, ImageType.Svg) and renders the chart to QuarterlySales.svg while also saving the workbook as QuarterlySales.xlsx.
// Keywords: Aspose.Cells SVG export | C# line chart to SVG | SvgImageOptions FitToViewPort | custom CSS prefix SVG | WOFF font embedding Aspose.Cells | ImageType.Svg .NET | scalable vector chart C# | Aspose.Cells chart rendering
// Common Searches: how to save Aspose.Cells chart as SVG in C# | SvgImageOptions settings for line charts Aspose.Cells | embed custom CSS prefix in SVG output Aspose.Cells | export Excel chart to scalable vector graphics .NET
// Developer Intent: Generate an SVG file from a worksheet chart using Aspose.Cells in a C# application.
// Use Cases: Display high‑resolution charts on web dashboards that scale without pixelation. | Prevent style conflicts by applying a unique CSS prefix to each exported SVG. | Ensure consistent typography across browsers by embedding a WOFF font directly in the SVG.
// AI Prompts: Write C# code that uses Aspose.Cells to export a bar chart to SVG with specific width and height. | Show how to configure SvgImageOptions to embed a TrueType font instead of WOFF. | Provide a loop that iterates through all charts in a workbook and saves each as an individual SVG file.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// The sample creates a workbook, fills cells A1:B4 with month and sales values, adds a line chart, sets SvgImageOptions (FitToViewPort, custom CSS prefix, WOFF font embedding, ImageType.Svg) and renders the chart to QuarterlySales.svg while also saving the workbook as QuarterlySales.xlsx.
class SvgChartGenerator
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Month");
        worksheet.Cells["B1"].PutValue("Sales");
        worksheet.Cells["A2"].PutValue("Jan");
        worksheet.Cells["B2"].PutValue(12000);
        worksheet.Cells["A3"].PutValue("Feb");
        worksheet.Cells["B3"].PutValue(15000);
        worksheet.Cells["A4"].PutValue("Mar");
        worksheet.Cells["B4"].PutValue(18000);

        // Add a line chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Line, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories
        chart.Title.Text = "Quarterly Sales";

        // Configure SVG rendering options
        SvgImageOptions svgOptions = new SvgImageOptions();
        svgOptions.FitToViewPort = true;                 // Fit SVG to viewport
        svgOptions.CssPrefix = "mychart-";               // Custom CSS prefix
        svgOptions.EmbeddedFontType = SvgEmbeddedFontType.Woff; // Embed WOFF font
        svgOptions.ImageType = ImageType.Svg;            // Ensure SVG output

        // Render the chart to an SVG file
        chart.ToImage("QuarterlySales.svg", svgOptions);

        // Optional: Save the workbook for reference
        workbook.Save("QuarterlySales.xlsx");
    }
}
