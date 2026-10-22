// Title: Export a chart to PDF with an 8.5 × 11 inch custom page size using Aspose.Cells in C#
// AI Prompts: Generate C# code that creates a worksheet, adds a chart, sets the worksheet PageSetup to an 8.5 × 11 inch custom paper size, and saves the workbook as a PDF with PdfSaveOptions. | Show how to configure PdfSaveOptions and PageSetup so that an Aspose.Cells chart fits on a single 8.5 × 11 inch PDF page in a .NET application.
// Common Searches: Aspose.Cells C# export chart to PDF with 8.5x11 page size | set worksheet page size 8.5 by 11 inches before PDF save Aspose.Cells | how to keep Excel chart on one PDF page using Aspose.Cells .NET | PdfSaveOptions OnePagePerSheet custom paper dimensions Aspose.Cells | C# Aspose.Cells chart PDF export custom paper size example
// Tags: chart export to PDF Aspose.Cells | custom paper size worksheet PageSetup | PdfSaveOptions OnePagePerSheet setting | Aspose.Cells set page dimensions C# | export Excel chart as 8.5x11 PDF

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Saving;

// The example creates a workbook, adds sample data and a column chart, sets the worksheet PageSetup to an 8.5 × 11 inch custom paper size, configures PdfSaveOptions to keep the chart on a single page, and saves the result as ChartExport.pdf.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIdx];
            chart.NSeries.Add("B2:B4", true);               // Values
            chart.NSeries.CategoryData = "A2:A4";           // Categories
            chart.Title.Text = "Sample Chart";

            // Configure PDF save options (default Letter size 8.5×11 inches)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true    // Keep the chart on a single page
            };

            // Export the workbook (containing the chart) to a PDF file
            workbook.Save("ChartExport.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
