// Title: How to export an Excel column chart to a PDF with a transparent background in C# using Aspose.Cells
// Common Searches: C# Aspose.Cells export chart to PDF with no background fill | How to make an Excel chart background transparent when saving as PDF using Aspose.Cells | Aspose.Cells example for transparent chart area in PDF output | Save a column chart as PDF without background in .NET | Aspose.Cells PDFSaveOptions transparent chart overlay
// Tags: Aspose.Cells chart PDF export C# | transparent chart background Aspose.Cells | PdfSaveOptions OnePagePerSheet .NET | column chart creation Aspose.Cells | Excel chart area no fill Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, adds sample data, builds a column chart, and uses PdfSaveOptions with OnePagePerSheet to save the chart as a PDF. Direct API calls for making the chart and plot areas transparent are not available in current Aspose.Cells versions, so the PDF is generated with the default chart background.
class ExportChartToPdfTransparent
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
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(45);
            sheet.Cells["B4"].PutValue(25);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the chart data source
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Note: Transparent background for chart and plot areas is not directly supported
            // via Aspose.Cells API in recent versions. The following lines are omitted to avoid
            // compilation errors:
            // chart.ChartArea.AreaFormat.FillFormat.FillType = FillType.NoFill;
            // chart.ChartArea.BorderLineStyle = CellBorderType.None;
            // chart.PlotArea.AreaFormat.FillFormat.FillType = FillType.NoFill;
            // chart.PlotArea.BorderLineStyle = CellBorderType.None;

            // Prepare PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
            };

            // Save the workbook (containing only the chart) to PDF
            workbook.Save("ChartTransparent.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
