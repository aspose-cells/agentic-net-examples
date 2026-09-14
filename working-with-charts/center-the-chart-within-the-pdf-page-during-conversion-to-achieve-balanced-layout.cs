// Title: How to center an Aspose.Cells chart on a PDF page using C# and the Chart.ToPdf method
// AI Prompts: Write C# code that creates a worksheet, builds a column chart, and exports it to a PDF with both horizontal and vertical alignment set to Center via the Chart.ToPdf method. | Show how to define custom page dimensions and apply PageLayoutAlignmentType.Center to position a chart in the middle of the PDF during conversion with Aspose.Cells.
// Common Searches: Aspose.Cells C# export chart to PDF centered on page | Chart.ToPdf alignment parameters example in .NET | How to set horizontal and vertical alignment for chart PDF export Aspose.Cells | Centering Excel chart in PDF using Aspose.Cells with custom page size | PageLayoutAlignmentType.Center usage for chart PDF conversion
// Tags: center alignment for chart PDF export | Aspose.Cells chart PDF conversion with custom page dimensions | using PageLayoutAlignmentType to align chart in PDF | C# chart export to PDF centered layout | Aspose.Cells PDF page size and chart positioning

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ChartPdfCenterExample
{
    // The example creates a workbook, adds sample data, inserts a column chart, and uses Chart.ToPdf to export the chart to a PDF named CenteredChart.pdf. It specifies an 8.5 × 11 inch page and applies PageLayoutAlignmentType.Center for both horizontal and vertical alignment, resulting in a chart that is centered on the PDF page.
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
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["A3"].PutValue("Bananas");
            sheet.Cells["A4"].PutValue("Cherries");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(40);
            sheet.Cells["B3"].PutValue(55);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Define desired PDF page size (in inches)
            float pageWidth = 8.5f;   // Standard Letter width
            float pageHeight = 11f;   // Standard Letter height

            // Export the chart to PDF with both horizontal and vertical alignment set to Center
            chart.ToPdf("CenteredChart.pdf", pageWidth, pageHeight,
                        PageLayoutAlignmentType.Center, PageLayoutAlignmentType.Center);

            Console.WriteLine("Chart has been exported to PDF with centered alignment.");
        }
    }
}
