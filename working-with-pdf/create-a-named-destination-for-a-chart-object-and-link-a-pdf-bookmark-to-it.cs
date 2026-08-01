// Title: Create a PDF bookmark linked to a chart via named destination using Aspose.Cells for .NET
// Description: This C# example shows how to build a workbook, add a column chart, define a named destination for the chart, create a PDF bookmark that points to that destination, and save the file with PdfSaveOptions so the bookmark jumps directly to the chart in the generated PDF.
// Keywords: Aspose.Cells PDF bookmark | named destination chart | C# Aspose.Cells PDF export | PdfSaveOptions bookmark | chart navigation PDF | Aspose.Cells example | PDF outline Aspose.Cells
// Common Searches: Aspose.Cells add PDF bookmark to chart | named destination for chart in PDF using Aspose.Cells | C# export chart with bookmark Aspose.Cells | how to link PDF bookmark to chart Aspose.Cells | PdfSaveOptions bookmark destination example
// Developer Intent: Add a PDF bookmark that navigates to a chart by using a named destination in Aspose.Cells for .NET.
// Use Cases: Generate a sales report where the PDF outline includes a "Sales Chart" entry that opens directly to the chart. | Create multi‑page PDFs with several charts, each accessible via its own bookmark for fast navigation. | Build interactive PDFs for presentations, allowing readers to jump to specific chart visualizations from the bookmark pane.
// AI Prompts: Write C# code with Aspose.Cells to define a named destination for a chart and add a PDF bookmark that links to it. | Show how to set PdfSaveOptions.Bookmark.DestinationName to reference a chart instead of a cell. | Provide an example that adds multiple chart bookmarks, each with a unique named destination, in a single PDF export.

using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// This C# example shows how to build a workbook, add a column chart, define a named destination for the chart, create a PDF bookmark that points to that destination, and save the file with PdfSaveOptions so the bookmark jumps directly to the chart in the generated PDF.
class Program
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
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(30);
        worksheet.Cells["B3"].PutValue(45);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data source for the chart
        chart.NSeries.Add("B2:B3", true);
        chart.NSeries.CategoryData = "A2:A3";

        // Assign a name to the chart (optional, helps identify the chart)
        chart.Name = "SalesChart";

        // Create a PDF bookmark entry that points to a named destination
        PdfBookmarkEntry bookmark = new PdfBookmarkEntry
        {
            Text = "Sales Chart",                 // Bookmark title shown in PDF viewer
            Destination = worksheet.Cells["A1"], // Anchor cell for the destination
            DestinationName = "ChartDest",        // Named destination identifier
            IsOpen = true                         // Expand the bookmark by default
        };

        // Configure PDF save options to include the bookmark
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Bookmark = bookmark
        };

        // Save the workbook as a PDF with the bookmark linked to the chart's destination
        workbook.Save("ChartWithBookmark.pdf", pdfOptions);
    }
}
