using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsChartBookmarkDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["A4"].PutValue("Cherry");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(45);
            sheet.Cells["B4"].PutValue(25);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Assign a name to the chart – this name will be used as a named destination
            chart.Name = "MyChartDestination";

            // Create a PDF bookmark entry that points to the chart via the named destination
            PdfBookmarkEntry bookmark = new PdfBookmarkEntry
            {
                Text = "Chart Bookmark",
                // Destination cell is required; it can be any cell on the sheet.
                Destination = sheet.Cells["A1"],
                // Set the named destination to the chart's name
                DestinationName = chart.Name,
                IsOpen = true
            };

            // Configure PDF save options with the bookmark
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Bookmark = bookmark
            };

            // Save the workbook as a PDF; the bookmark will link to the chart's named destination
            workbook.Save("ChartWithBookmark.pdf", pdfOptions);

            Console.WriteLine("PDF with chart and linked bookmark created successfully.");
        }
    }
}