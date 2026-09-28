// Title: How to add a PDF bookmark that points to cell A1 on a chart sheet using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that creates a PdfBookmarkEntry for cell A1, attaches it to a chart sheet's Bookmark collection, and saves the workbook as a PDF. | Show the steps to link a PDF bookmark to cell A1 on a chart sheet in Aspose.Cells before exporting, and explain the current limitation for chart‑sheet bookmarks.
// Common Searches: Aspose.Cells C# add PDF bookmark to chart sheet cell A1 | Create PdfBookmarkEntry for chart sheet using Aspose.Cells .NET | Export chart sheet to PDF with cell bookmark Aspose.Cells example | PdfBookmarkEntry not working on chart sheets Aspose.Cells limitation | How to assign a PDF bookmark to a chart sheet in Aspose.Cells for .NET
// Tags: aspnet cells pdf bookmark chart sheet | c# pdf bookmark entry aspnet cells | export chart sheet to pdf aspnet cells | pdf outline from worksheet cell aspnet cells | aspnet cells chart sheet bookmark limitation

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example demonstrates creating a workbook, adding a worksheet that hosts a column chart, populating cells A1‑B4, and attempting to assign a PdfBookmarkEntry for cell A1 to the chart sheet's Bookmark before saving as PDF. It also notes that PDF bookmarks are not supported on chart sheets in the current Aspose.Cells version.
class PdfBookmarkExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a regular worksheet that will host the chart (chart sheets are not supported in this version)
            int chartSheetIndex = workbook.Worksheets.Add();
            Worksheet chartSheet = workbook.Worksheets[chartSheetIndex];
            chartSheet.Name = "MyChartSheet";

            // Add a simple column chart to the worksheet
            int chartIndex = chartSheet.Charts.Add(ChartType.Column, 0, 0, 15, 20);
            Chart chart = chartSheet.Charts[chartIndex];

            // Optionally configure the chart (e.g., add sample data)
            // Here we add dummy data to demonstrate the chart rendering
            chartSheet.Cells["A1"].PutValue("Category");
            chartSheet.Cells["B1"].PutValue("Value");
            chartSheet.Cells["A2"].PutValue("Item 1");
            chartSheet.Cells["B2"].PutValue(10);
            chartSheet.Cells["A3"].PutValue("Item 2");
            chartSheet.Cells["B3"].PutValue(20);
            chartSheet.Cells["A4"].PutValue("Item 3");
            chartSheet.Cells["B4"].PutValue(30);

            // Set the chart data source
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Save the workbook as PDF (bookmarks are not supported in this version)
            workbook.Save("ChartSheetWithBookmark.pdf", SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
