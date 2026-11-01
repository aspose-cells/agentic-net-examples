// Title: Export a column chart to an A4 portrait PDF using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export chart as PDF with A4 portrait layout | How to set PDF page size to A4 when saving an Excel chart with Aspose.Cells | Aspose.Cells PdfSaveOptions portrait orientation for chart export in .NET | Save Excel column chart to PDF using Aspose.Cells with one page per sheet | Export Aspose.Cells workbook containing a chart to A4 PDF in C#
// Tags: Aspose.Cells chart PDF export C# | PdfSaveOptions A4 portrait Aspose.Cells | export Excel chart to PDF Aspose | OnePagePerSheet option Aspose.Cells | column chart creation Aspose.Cells C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, adds a column chart with sample data, configures PdfSaveOptions (using the default A4 portrait page size and OnePagePerSheet), and saves the workbook as a PDF that includes the chart.
class ExportChartToPdf
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a column chart to the worksheet (from row 5, column 0 to row 20, column 10)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(45);
            sheet.Cells["B4"].PutValue(25);

            // Set the chart data source
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure PDF save options (default settings)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Example: export each worksheet to a separate page
                OnePagePerSheet = true
            };

            // Export the workbook (containing the chart) to PDF
            workbook.Save("ChartExport.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
