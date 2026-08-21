// Title: C# – Add a Diagonal WordArt‑Style Watermark Behind a Chart Plot Area When Exporting to PDF with Aspose.Cells
// Description: Demonstrates how to create a workbook, insert a column chart, define a large bold semi‑transparent font, build a RenderingWatermark with the text "CONFIDENTIAL" rotated 30°, set it as a background element, and apply the watermark via PdfSaveOptions so the watermark appears behind the chart plot area in the generated PDF.
// Keywords: Aspose.Cells | C# chart watermark | PDF export watermark | RenderingWatermark | diagonal watermark | WordArt style watermark | semi transparent watermark | background chart watermark | plot area watermark | Aspose.Cells for .NET
// Common Searches: Aspose.Cells add rotated watermark to chart PDF | C# diagonal WordArt watermark behind chart | RenderingWatermark example for chart export | How to place a background watermark on a chart in Aspose.Cells | Export chart to PDF with watermark Aspose.Cells
// Developer Intent: Place a diagonal, semi‑transparent WordArt‑style watermark behind the plot area of a chart when saving the workbook as a PDF using Aspose.Cells for .NET.
// Use Cases: Create confidential PDF reports where charts carry a diagonal "CONFIDENTIAL" overlay. | Brand exported chart PDFs with a rotated company logo or slogan as a background watermark. | Mark draft chart PDFs with a semi‑transparent, angled watermark to indicate they are not final.
// AI Prompts: Generate C# code that adds a 30° rotated RenderingWatermark behind a chart and saves the workbook as PDF with Aspose.Cells. | Explain how to customize font, color, opacity, and alignment of a RenderingWatermark for chart PDFs in Aspose.Cells. | Show how to ensure a watermark appears behind the plot area but in front of chart series when exporting to PDF.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// Demonstrates how to create a workbook, insert a column chart, define a large bold semi‑transparent font, build a RenderingWatermark with the text "CONFIDENTIAL" rotated 30°, set it as a background element, and apply the watermark via PdfSaveOptions so the watermark appears behind the chart plot area in the generated PDF.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a column chart covering the data range
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Create a rendering font for the watermark (WordArt‑like appearance)
        RenderingFont watermarkFont = new RenderingFont("Arial", 72)
        {
            Bold = true,
            Color = Color.FromArgb(128, 0, 0, 255) // semi‑transparent blue
        };

        // Create the watermark with the desired text and set rotation to 30°
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", watermarkFont)
        {
            Rotation = 30,               // diagonal effect
            Opacity = 0.3f,              // semi‑transparent
            IsBackground = true,         // placed behind content
            HAlignment = TextAlignmentType.Center,
            VAlignment = TextAlignmentType.Center
        };

        // Configure PDF save options to include the watermark
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Watermark = watermark
        };

        // Save the workbook as a PDF; the watermark appears over the chart plot area
        workbook.Save("ChartWatermark.pdf", pdfOptions);
    }
}
