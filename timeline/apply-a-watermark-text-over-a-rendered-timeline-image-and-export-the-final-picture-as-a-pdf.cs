// Title: Add a semi‑transparent diagonal CONFIDENTIAL text watermark to a timeline line chart and export it as PDF using Aspose.Cells for .NET
// AI Prompts: Generate C# code that builds a line chart from date/value cells, applies a 45‑degree semi‑transparent text watermark with RenderingWatermark, and saves the workbook as a PDF using Aspose.Cells. | Show how to configure PdfSaveOptions with a RenderingWatermark to overlay a centered, rotated watermark on a rendered chart image during PDF export. | Explain how to set watermark font, opacity, rotation, and background flag when exporting an Excel worksheet containing a timeline chart to PDF with Aspose.Cells.
// Common Searches: aspocells how to add a diagonal text watermark to a chart when exporting to PDF in C# | c# render timeline chart as PDF with confidential watermark using Aspose.Cells | set watermark opacity and rotation in PdfSaveOptions Aspose.Cells .NET | export Excel line chart with watermark overlay to PDF using Aspose.Cells
// Tags: Aspose.Cells RenderingWatermark PDF generation | timeline chart watermarking C# | PdfSaveOptions watermark settings Aspose.Cells | export Excel chart with diagonal text overlay | semi‑transparent rotated watermark .NET

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// // Creates a workbook with date/value data, adds a line chart as a timeline, defines a semi‑transparent 45° "CONFIDENTIAL" watermark via RenderingWatermark, assigns it to PdfSaveOptions, and saves the workbook as a PDF where the chart image includes the watermark.
class TimelineWatermarkPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data that represents a timeline (date vs value)
        sheet.Cells["A1"].PutValue("Date");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue(new DateTime(2023, 2, 1));
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue(new DateTime(2023, 3, 1));
        sheet.Cells["B4"].PutValue(15);

        // Add a line chart that will be rendered as a timeline image
        int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 15);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Dates as categories
        chart.Title.Text = "Project Timeline";

        // Create a text watermark using RenderingWatermark
        RenderingFont watermarkFont = new RenderingFont("Arial", 48)
        {
            Color = Color.LightGray,
            Bold = true,
            Italic = true
        };

        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", watermarkFont)
        {
            HAlignment = TextAlignmentType.Center,
            VAlignment = TextAlignmentType.Center,
            Rotation = 45f,
            Opacity = 0.3f,
            IsBackground = true
        };

        // Set PDF save options with the watermark
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Watermark = watermark
        };

        // Save the workbook as a PDF; the chart is rendered as an image and the watermark is applied over it
        workbook.Save("TimelineWithWatermark.pdf", pdfOptions);
    }
}
