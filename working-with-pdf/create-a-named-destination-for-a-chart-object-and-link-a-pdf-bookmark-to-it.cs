// Title: Add a PDF Bookmark to a Chart Using a Named Destination with Aspose.Cells for .NET
// Description: Learn how to create a workbook, insert a column chart, assign a name to the chart, define a PdfBookmarkEntry that points to a cell near the chart as a named destination, and export the workbook to PDF with PdfSaveOptions so the generated PDF contains a clickable bookmark that jumps directly to the chart.
// Keywords: Aspose.Cells PDF bookmark | named destination chart | C# Aspose.Cells export PDF | PdfSaveOptions chart bookmark | Aspose.Cells chart to PDF | PDF outline Aspose.Cells | C# create PDF bookmark
// Common Searches: how to add a PDF bookmark to a chart in Aspose.Cells | Aspose.Cells named destination for chart export | C# export chart with PDF bookmark using Aspose.Cells | set PDF bookmark destination to a chart cell Aspose.Cells | Aspose.Cells PdfSaveOptions bookmark example
// Developer Intent: Create a PDF bookmark that links to a named destination representing a chart when exporting a workbook with Aspose.Cells for .NET.
// Use Cases: Generate a sales report PDF where each chart has its own bookmark for instant navigation. | Build an interactive PDF with multiple chart bookmarks that jump to the corresponding visualizations. | Define a reusable named destination for a chart so several bookmarks or internal links can reference the same chart area.
// AI Prompts: Show me how to add multiple PDF bookmarks for several charts in a workbook using Aspose.Cells for .NET. | Explain how to set a named destination for a chart and reference it from a PDF bookmark in C#. | Provide code to create a PDF outline with bookmarks that point to different chart objects in the same workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsChartBookmarkDemo
{
    // Learn how to create a workbook, insert a column chart, assign a name to the chart, define a PdfBookmarkEntry that points to a cell near the chart as a named destination, and export the workbook to PDF with PdfSaveOptions so the generated PDF contains a clickable bookmark that jumps directly to the chart.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["A4"].PutValue("Cherry");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(45);
            sheet.Cells["B4"].PutValue(25);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";
            chart.Title.Text = "Fruit Sales";

            // Assign a name to the chart (optional, useful for reference)
            chart.Name = "FruitSalesChart";

            // Create a PDF bookmark entry that points to the cell where the chart starts (A5)
            PdfBookmarkEntry chartBookmark = new PdfBookmarkEntry
            {
                Text = "Chart Bookmark",
                Destination = sheet.Cells["A5"],          // Cell near the chart
                DestinationName = "ChartDestination",    // Named destination
                IsOpen = true
            };

            // Configure PDF save options with the bookmark
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Bookmark = chartBookmark
            };

            // Save the workbook as a PDF; the bookmark will link to the named destination
            workbook.Save("ChartWithBookmark.pdf", pdfOptions);
        }
    }
}
