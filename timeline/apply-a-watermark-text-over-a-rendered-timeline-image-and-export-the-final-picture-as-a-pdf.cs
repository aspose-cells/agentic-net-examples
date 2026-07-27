// Title: Add a CONFIDENTIAL text watermark to a timeline chart PDF with Aspose.Cells for .NET (C#)
// Description: This C# example creates a workbook, fills it with date‑event data, builds a line chart that visualizes a project timeline, defines a large red italic watermark rotated 45°, sets 30 % opacity, and applies the watermark via PdfSaveOptions so the final PDF shows the watermark over the rendered chart image.
// Keywords: Aspose.Cells C# watermark PDF | RenderingWatermark example | timeline chart export | PdfSaveOptions watermark | diagonal text watermark .NET | confidential watermark Aspose.Cells | chart to PDF Aspose.Cells | Aspose.Cells rendering options
// Common Searches: how to add a rotated text watermark to a PDF generated from Aspose.Cells | Aspose.Cells C# export timeline chart as PDF with watermark | RenderingWatermark usage in Aspose.Cells .NET | apply CONFIDENTIAL watermark to chart PDF Aspose | set opacity and rotation for Aspose.Cells PDF watermark
// Developer Intent: Overlay a semi‑transparent, rotated text watermark on a rendered timeline chart and save the result as a PDF.
// Use Cases: Distribute project timelines with a CONFIDENTIAL label to protect sensitive information. | Include a corporate watermark on sales or financial charts exported to PDF for compliance. | Automate weekly reporting where every chart page carries a draft or internal‑use watermark.
// AI Prompts: Show how to change the opacity and rotation angle of a RenderingWatermark before exporting to PDF with Aspose.Cells. | Provide a sample that adds both text and image watermarks to a workbook and saves it as a PDF. | Explain how to restrict a watermark to the chart area instead of the whole PDF page using Aspose.Cells rendering settings.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// This C# example creates a workbook, fills it with date‑event data, builds a line chart that visualizes a project timeline, defines a large red italic watermark rotated 45°, sets 30 % opacity, and applies the watermark via PdfSaveOptions so the final PDF shows the watermark over the rendered chart image.
class TimelineWatermarkPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate data that will be used for a simple timeline chart
        sheet.Cells["A1"].PutValue("Date");
        sheet.Cells["B1"].PutValue("Event");
        sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
        sheet.Cells["B2"].PutValue("Start");
        sheet.Cells["A3"].PutValue(new DateTime(2023, 2, 15));
        sheet.Cells["B3"].PutValue("Milestone");
        sheet.Cells["A4"].PutValue(new DateTime(2023, 4, 30));
        sheet.Cells["B4"].PutValue("Finish");

        // Add a line chart that visualizes the timeline
        int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 15);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Values (events)
        chart.NSeries.CategoryData = "A2:A4";      // Dates as categories
        chart.Title.Text = "Project Timeline";

        // Create a rendering font for the watermark text
        RenderingFont watermarkFont = new RenderingFont("Arial", 48)
        {
            Color = Color.Red,
            Bold = true,
            Italic = true
        };

        // Create a text watermark and configure its appearance
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", watermarkFont)
        {
            HAlignment = TextAlignmentType.Center,
            VAlignment = TextAlignmentType.Center,
            Rotation = 45f,
            Opacity = 0.3f,
            IsBackground = true
        };

        // Set PDF save options to include the watermark
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Watermark = watermark
        };

        // Save the workbook as a PDF; the watermark will be applied over the rendered timeline image
        workbook.Save("TimelineWithWatermark.pdf", pdfOptions);
    }
}
