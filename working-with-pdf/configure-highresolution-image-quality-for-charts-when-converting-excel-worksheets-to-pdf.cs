// Title: Set High‑Resolution Chart Images When Converting Excel to PDF with Aspose.Cells for .NET
// Description: This example creates a workbook, adds a column chart, and uses PdfSaveOptions.SetImageResample to export the sheet as a PDF with 300 PPI chart images and maximum JPEG quality, preserving crisp detail for print‑ready documents.
// Keywords: Aspose.Cells | .NET | C# | PdfSaveOptions | SetImageResample | high DPI chart PDF | 300 PPI Excel to PDF | chart image quality | PDF optimization | Aspose.Cells USA
// Common Searches: Aspose.Cells set image resample 300 DPI | high resolution chart PDF Aspose.Cells C# | increase chart quality when saving Excel as PDF | PdfSaveOptions image quality settings | export Excel chart to PDF with high DPI
// Developer Intent: Configure PDF export options so that Excel chart images are rendered at high resolution and maximum JPEG quality.
// Use Cases: Produce printable sales reports where column charts remain sharp at 300 PPI. | Generate client‑ready financial dashboards with high‑DPI charts in PDF format. | Batch‑export workbooks containing multiple charts while preserving image clarity for professional printing.
// AI Prompts: Show how to change SetImageResample to 600 DPI and 90% JPEG quality for PDF conversion in Aspose.Cells. | Provide a C# snippet that saves an Excel file to PDF with high‑resolution charts and disables image compression. | Explain how PdfOptimizationType affects chart rendering quality in Aspose.Cells PDF output.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// This example creates a workbook, adds a column chart, and uses PdfSaveOptions.SetImageResample to export the sheet as a PDF with 300 PPI chart images and maximum JPEG quality, preserving crisp detail for print‑ready documents.
class HighResolutionChartPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["A3"].PutValue("Banana");
        worksheet.Cells["A4"].PutValue("Cherry");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(120);
        worksheet.Cells["B3"].PutValue(80);
        worksheet.Cells["B4"].PutValue(150);

        // Add a column chart and bind it to the data range
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";
        chart.Title.Text = "Fruit Sales";

        // Configure PDF save options to render chart images at high resolution
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        // Desired PPI (e.g., 300) and JPEG quality (0‑100). 300 PPI gives high‑quality vector‑like rendering.
        pdfOptions.SetImageResample(300, 100);
        // Keep standard print quality (optional, can be omitted)
        pdfOptions.OptimizationType = PdfOptimizationType.Standard;

        // Save the workbook as PDF using the configured options
        workbook.Save("HighResolutionChart.pdf", pdfOptions);
    }
}
