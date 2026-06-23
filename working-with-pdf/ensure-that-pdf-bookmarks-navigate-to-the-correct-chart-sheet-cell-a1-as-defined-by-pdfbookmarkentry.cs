using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a worksheet that will contain the chart
            int sheetIndex = workbook.Worksheets.Add();
            Worksheet chartSheet = workbook.Worksheets[sheetIndex];
            chartSheet.Name = "ChartSheet";

            // Put an identifier in cell A1 (bookmark destination)
            chartSheet.Cells["A1"].PutValue("Chart Sheet Start");

            // Populate sample data for the chart
            chartSheet.Cells["B1"].PutValue("Category");
            chartSheet.Cells["C1"].PutValue("Value");
            chartSheet.Cells["B2"].PutValue("A");
            chartSheet.Cells["C2"].PutValue(10);
            chartSheet.Cells["B3"].PutValue("B");
            chartSheet.Cells["C3"].PutValue(20);

            // Add a column chart to the worksheet
            int chartIdx = chartSheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = chartSheet.Charts[chartIdx];
            chart.NSeries.Add("C2:C3", true);          // Values
            chart.NSeries.CategoryData = "B2:B3";      // Categories

            // Create a PDF bookmark that points to cell A1 of the chart sheet
            PdfBookmarkEntry bookmark = new PdfBookmarkEntry
            {
                Text = "Chart Sheet Start",
                Destination = chartSheet.Cells["A1"],
                IsOpen = true
            };

            // Set PDF save options with the bookmark
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Bookmark = bookmark
            };

            // Save the workbook as a PDF; the bookmark will navigate to A1 of the chart sheet
            workbook.Save("ChartSheetBookmark.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}