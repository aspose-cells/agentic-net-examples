// Title: Export Aspose.Cells Chart to PDF (A4 portrait) with centered alignment in C#
// Description: Creates a workbook, adds a column chart, sets the chart's page setup to A4 portrait, and uses Chart.ToPdf to generate a centered PDF (8.27 × 11.69 in) named ChartA4Portrait.pdf.
// Keywords: Aspose.Cells chart to PDF | C# export chart PDF | A4 portrait PDF Aspose.Cells | Chart.ToPdf alignment | set paper size Aspose.Cells
// Common Searches: Aspose.Cells export chart to PDF A4 portrait C# | How to set page size for chart PDF in Aspose.Cells | C# chart ToPdf with custom width and height | center chart on PDF using Aspose.Cells | Aspose.Cells chart page setup orientation
// Developer Intent: Generate a PDF file from an Aspose.Cells chart using A4 portrait dimensions and center the chart on the page.
// Use Cases: Produce printable sales charts for reports that must fit an A4 page. | Automate batch conversion of workbook charts to A4 portrait PDFs for archival. | Create PDF handouts of dashboards where each chart needs consistent page size and alignment.
// AI Prompts: Write C# code with Aspose.Cells to export a chart to an A4 portrait PDF and center it. | Show how to configure PaperSize and Orientation for a chart before calling ToPdf. | Explain the width and height parameters of Chart.ToPdf for different page formats.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsChartPdfExport
{
    // Creates a workbook, adds a column chart, sets the chart's page setup to A4 portrait, and uses Chart.ToPdf to generate a centered PDF (8.27 × 11.69 in) named ChartA4Portrait.pdf.
    public class ExportChartToPdfA4Portrait
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

                // Configure page setup for A4 portrait
                chart.PageSetup.PaperSize = PaperSizeType.PaperA4;
                chart.PageSetup.Orientation = PageOrientationType.Portrait;

                // Export the chart to PDF with A4 size (8.27 x 11.69 inches) and centered alignment
                chart.ToPdf(
                    "ChartA4Portrait.pdf",
                    8.27f,               // page width in inches (A4 width)
                    11.69f,              // page height in inches (A4 height)
                    PageLayoutAlignmentType.Center, // horizontal alignment
                    PageLayoutAlignmentType.Center  // vertical alignment
                );

                Console.WriteLine("Chart exported to PDF with A4 portrait orientation successfully.");
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
            ExportChartToPdfA4Portrait.Run();
        }
    }
}
