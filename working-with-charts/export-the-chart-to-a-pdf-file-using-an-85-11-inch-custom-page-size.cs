// Title: Export a worksheet chart to a PDF with letter‑size (8.5 × 11 in) using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export chart to PDF with 8.5x11 page size | How to set PDF page dimensions to letter size when saving a chart in Aspose.Cells | Saving an Excel column chart as PDF using PdfSaveOptions in .NET | Aspose.Cells one page per sheet PDF export example for charts | Configure page setup for PDF output of a chart with Aspose.Cells C#
// Tags: Aspose.Cells chart PDF export | C# PdfSaveOptions OnePagePerSheet | letter size PDF page setup Aspose.Cells | export column chart to PDF | worksheet chart to PDF Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The example creates a workbook, adds a column chart sourced from cells A1:A3, configures PdfSaveOptions to place one page per sheet, and saves the workbook (including the chart) as a PDF file named Chart.pdf, producing a letter‑size (8.5 × 11 in) document.
class ExportChartToPdf
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data for the chart
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);

            // Add a column chart and set its data source
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 7);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("A1:A3", true);
            chart.Title.Text = "Sample Chart";

            // Optional: configure PDF save options (one page per sheet)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
            };

            // Ensure the output directory exists
            string outputPath = "Chart.pdf";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (which contains the chart) as a PDF file
            workbook.Save(outputPath, pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
