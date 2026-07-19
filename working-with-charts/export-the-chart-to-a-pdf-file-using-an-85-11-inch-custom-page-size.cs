// Title: Export Aspose.Cells Chart to PDF with Custom 8.5×11 in Page Size and Center Alignment (C#)
// Description: Shows how to create a workbook, add sample data, insert a column chart, and call Chart.ToPdf to generate ChartCustomSize.pdf using an 8.5 × 11 in page with both horizontal and vertical centering.
// Keywords: Aspose.Cells | Chart.ToPdf | C# | .NET | PDF export | custom page size | letter size PDF | center alignment | chart export | Excel chart to PDF
// Common Searches: Aspose.Cells export chart to PDF with specific dimensions | Chart.ToPdf custom width height example C# | center chart on PDF page using Aspose.Cells | set page size for chart PDF export Aspose.Cells .NET | how to create letter‑size PDF from an Excel chart
// Developer Intent: Export a chart to a PDF file using an 8.5×11 in page and centered layout.
// Use Cases: Produce printable reports where charts must fit standard letter pages. | Automate PDF generation for marketing materials that require centered graphics. | Create documentation with consistently sized and positioned charts without manual editing.
// AI Prompts: Generate C# code with Aspose.Cells that saves a chart as a PDF using an 8.5×11 in page and centers it. | Explain each parameter of Chart.ToPdf when defining custom page dimensions and alignment. | Provide troubleshooting steps if the exported PDF chart appears off‑center or blank.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartExport
{
    // Shows how to create a workbook, add sample data, insert a column chart, and call Chart.ToPdf to generate ChartCustomSize.pdf using an 8.5 × 11 in page with both horizontal and vertical centering.
    public class ExportChartToPdf
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("Fruits");
                worksheet.Cells["A3"].PutValue("Vegetables");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(50);
                worksheet.Cells["B3"].PutValue(30);

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B3", true);
                chart.NSeries.CategoryData = "A2:A3";

                // Export the chart to a PDF file with a custom page size of 8.5 × 11 inches.
                chart.ToPdf(
                    "ChartCustomSize.pdf",          // output file name
                    8.5f,                          // desired page width in inches
                    11f,                           // desired page height in inches
                    PageLayoutAlignmentType.Center, // horizontal alignment
                    PageLayoutAlignmentType.Center  // vertical alignment
                );

                Console.WriteLine("Chart exported to PDF with custom page size successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportChartToPdf.Run();
        }
    }
}
