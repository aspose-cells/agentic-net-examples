// Title: Center a Chart on a PDF Page with Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, add a column chart, set US Letter page dimensions, and export the chart to PDF using Chart.ToPdf with PageLayoutAlignmentType.Center for both horizontal and vertical alignment, producing a centered chart layout.
// Keywords: Aspose.Cells | C# chart to PDF | center chart PDF | PageLayoutAlignmentType.Center | US Letter PDF export | chart alignment Aspose.Cells | .NET PDF chart export
// Common Searches: center chart Aspose.Cells PDF export | Aspose.Cells align chart on PDF page | Chart.ToPdf horizontal vertical centering C# | set page size and center chart in PDF using Aspose.Cells | how to position chart in the middle of a PDF with Aspose
// Developer Intent: Export a chart to a PDF file and position it at the center of the page.
// Use Cases: Generate a sales performance PDF where the chart is centered for a clean presentation. | Create marketing flyers in PDF format that feature a centrally placed chart. | Automate financial reports that embed a centered chart for easy visual analysis.
// AI Prompts: Modify the example to add custom margins while keeping the chart centered. | Provide code to center multiple charts on separate PDF pages using Aspose.Cells. | Explain the effect of PageLayoutAlignmentType values when exporting charts to PDF.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartCenterPdf
{
    // Shows how to create a workbook, add a column chart, set US Letter page dimensions, and export the chart to PDF using Chart.ToPdf with PageLayoutAlignmentType.Center for both horizontal and vertical alignment, producing a centered chart layout.
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

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Define desired PDF page size (in inches)
            float pageWidth = 8.5f;   // Standard US Letter width
            float pageHeight = 11f;   // Standard US Letter height

            // Export the chart to PDF, centering it both horizontally and vertically
            chart.ToPdf("CenteredChart.pdf", pageWidth, pageHeight,
                        PageLayoutAlignmentType.Center, PageLayoutAlignmentType.Center);

            Console.WriteLine("Chart has been exported to 'CenteredChart.pdf' with centered alignment.");
        }
    }
}
