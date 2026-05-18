using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Charts;

class TimelineWatermarkPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data representing a timeline (date vs. value)
        sheet.Cells["A1"].PutValue("Date");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue(new DateTime(2023, 2, 1));
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue(new DateTime(2023, 3, 1));
        sheet.Cells["B4"].PutValue(15);

        // Add a line chart that will act as the timeline visual
        int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 15);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Dates
        chart.Title.Text = "Project Timeline";

        // Create a rendering font for the watermark text
        RenderingFont watermarkFont = new RenderingFont("Arial", 48)
        {
            Color = Color.Red,
            Bold = true,
            Italic = true
        };

        // Instantiate a text watermark and configure its appearance
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", watermarkFont)
        {
            HAlignment = TextAlignmentType.Center,
            VAlignment = TextAlignmentType.Center,
            Rotation = 45f,          // Rotate 45 degrees
            Opacity = 0.3f,          // Semi‑transparent
            IsBackground = true      // Place behind page contents
        };

        // Set PDF save options to include the watermark
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Watermark = watermark
        };

        // Save the workbook as a PDF; the chart (timeline) will be rendered with the watermark
        workbook.Save("TimelineWithWatermark.pdf", pdfOptions);
    }
}