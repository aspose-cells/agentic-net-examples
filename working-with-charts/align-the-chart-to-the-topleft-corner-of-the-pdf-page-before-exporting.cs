// Title: Export Aspose.Cells Chart to PDF Aligned Top‑Left (C#)
// Description: Demonstrates how to create a workbook, add a column chart, and export the chart to a PDF positioned at the top‑left corner of the page using Chart.ToPdf with custom page dimensions and PageLayoutAlignmentType.Left/Top.
// Keywords: Aspose.Cells | Chart.ToPdf | C# | PDF export | chart alignment | top left | PageLayoutAlignmentType | page size | column chart | Aspose.Cells PDF positioning
// Common Searches: Aspose.Cells align chart top left PDF | Chart.ToPdf horizontal vertical alignment C# | export chart to PDF with custom page size Aspose | position chart at page origin Aspose.Cells | C# code for chart PDF alignment Aspose
// Developer Intent: Export a chart to PDF and place it at the page's top‑left corner.
// Use Cases: Generate single‑chart PDFs where the chart starts at the page origin for clean header layouts. | Batch‑convert multiple worksheets' charts to PDFs with consistent top‑left positioning and fixed page dimensions. | Create printable reports where charts need to align with other page elements placed manually.
// AI Prompts: Write C# code that uses Aspose.Cells to export a chart to PDF with top‑left alignment and a custom page size. | Explain how PageLayoutAlignmentType.Left and PageLayoutAlignmentType.Top affect chart placement in Chart.ToPdf. | Show how to resize a chart after aligning it to the top‑left corner during PDF export with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook, add a column chart, and export the chart to a PDF positioned at the top‑left corner of the page using Chart.ToPdf with custom page dimensions and PageLayoutAlignmentType.Left/Top.
class AlignChartTopLeftPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data source for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";
        chart.Title.Text = "Sample Chart";

        // Export the chart to PDF, aligning it to the top‑left corner of the page
        chart.ToPdf(
            "ChartTopLeft.pdf",          // output file
            8.5f,                        // page width in inches
            11f,                         // page height in inches
            PageLayoutAlignmentType.Left, // horizontal alignment
            PageLayoutAlignmentType.Top   // vertical alignment
        );
    }
}
