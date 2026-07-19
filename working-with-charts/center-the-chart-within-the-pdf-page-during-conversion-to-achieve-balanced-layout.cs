// Title: Center a Chart on a PDF Page with Aspose.Cells C# (US Letter)
// Description: Demonstrates how to export a column chart to a PDF and position it at the exact center of an US‑Letter page using the Chart.ToPdf overload with PageLayoutAlignmentType.Center for both horizontal and vertical alignment.
// Keywords: Aspose.Cells chart PDF center | Chart.ToPdf alignment C# | center chart on PDF page | PageLayoutAlignmentType.Center | US Letter PDF Aspose.Cells | C# export chart to PDF | Aspose.Cells chart layout
// Common Searches: center chart when exporting to PDF Aspose.Cells C# | Chart.ToPdf horizontal vertical alignment example | Aspose.Cells PDF page size and chart positioning | C# code to place chart in middle of PDF page | Aspose.Cells chart layout US Letter PDF
// Developer Intent: Export a chart to PDF and have it automatically centered on the page.
// Use Cases: Professional sales reports where the chart must be centered for visual balance. | Printable marketing flyers that require a centrally positioned chart to highlight metrics. | Automated PDF dashboards that maintain consistent chart alignment across multiple pages.
// AI Prompts: Generate C# code using Aspose.Cells to export a chart to a PDF with custom page dimensions and center alignment both horizontally and vertically. | Explain the effect of PageLayoutAlignmentType.Center in the Chart.ToPdf method and how to adapt it for A4, Letter, or custom paper sizes. | Show how to center several charts on separate PDF pages within a single workbook using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartCenterPdf
{
    // Demonstrates how to export a column chart to a PDF and position it at the exact center of an US‑Letter page using the Chart.ToPdf overload with PageLayoutAlignmentType.Center for both horizontal and vertical alignment.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
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
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Define PDF page size (in inches) and center alignment
            float pageWidth = 8.5f;   // Standard US Letter width
            float pageHeight = 11f;   // Standard US Letter height

            // Export the chart to PDF with centered alignment both horizontally and vertically
            chart.ToPdf("CenteredChart.pdf", pageWidth, pageHeight,
                        PageLayoutAlignmentType.Center, PageLayoutAlignmentType.Center);

            Console.WriteLine("Chart has been exported to PDF with centered layout.");
        }
    }
}
