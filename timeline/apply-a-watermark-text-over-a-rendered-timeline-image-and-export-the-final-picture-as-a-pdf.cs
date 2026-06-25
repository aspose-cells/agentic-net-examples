using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsWatermarkDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Populate sample data that will be used for a timeline (line) chart
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue(new DateTime(2023, 2, 1));
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["A4"].PutValue(new DateTime(2023, 3, 1));
            sheet.Cells["B4"].PutValue(130);
            sheet.Cells["A5"].PutValue(new DateTime(2023, 4, 1));
            sheet.Cells["B5"].PutValue(170);

            // 3. Add a line chart (used as a timeline) to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Line, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B5", true);          // Values
            chart.NSeries.CategoryData = "A2:A5";      // Dates as categories
            chart.Title.Text = "Project Timeline";

            // 4. Create a font for the watermark text
            RenderingFont watermarkFont = new RenderingFont("Arial", 48)
            {
                Bold = true,
                Color = Color.FromArgb(128, 0, 0, 255) // Semi‑transparent blue
            };

            // 5. Create the text watermark and configure its appearance
            RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", watermarkFont)
            {
                HAlignment = TextAlignmentType.Center,
                VAlignment = TextAlignmentType.Center,
                Rotation = 45f,
                Opacity = 0.3f,
                IsBackground = true
            };

            // 6. Set PDF save options with the watermark
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Watermark = watermark
            };

            // 7. Save the workbook as a PDF; the chart (timeline) will be rendered
            //    and the watermark will be applied over the whole page.
            string outputPath = "TimelineWithWatermark.pdf";
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF saved to: {outputPath}");
        }
    }
}