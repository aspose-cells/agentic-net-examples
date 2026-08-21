// Title: Export Scatter Chart to PDF with Original Size using Aspose.Cells for .NET
// Description: Creates a workbook, adds X/Y data, builds a scatter chart, reads its pixel dimensions, converts them to inches based on system DPI, and saves the chart to a PDF while preserving width, height, and centering on the page.
// Keywords: Aspose.Cells export scatter chart PDF | scatter chart to PDF .NET | preserve chart dimensions Aspose | chart GetActualSize pixels | convert pixels to inches DPI | Chart.ToPdf method | C# Aspose.Cells PDF export | center chart on PDF page
// Common Searches: export scatter chart to PDF Aspose.Cells | how to keep original chart size when saving as PDF | Aspose.Cells GetActualSize example | convert chart size from pixels to inches | center chart on PDF using Aspose.Cells
// Developer Intent: Generate a PDF file that contains a scatter chart rendered at the same size and resolution as it appears in Excel.
// Use Cases: Produce scientific reports where chart dimensions must match the on‑screen layout. | Create printable PDFs that replicate the exact worksheet design for regulatory submissions. | Automate batch conversion of multiple Excel scatter charts to PDFs while maintaining DPI consistency.
// AI Prompts: Write C# code with Aspose.Cells to export a scatter chart to PDF, preserving its original pixel dimensions and centering it. | Show how to retrieve a chart's actual size in pixels and convert the values to inches using the system DPI for PDF export. | Explain how to adjust DPI or scaling factors when exporting charts to PDF with Aspose.Cells.

using System;
using System.Drawing;                     // For Size
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds X/Y data, builds a scatter chart, reads its pixel dimensions, converts them to inches based on system DPI, and saves the chart to a PDF while preserving width, height, and centering on the page.
class ExportScatterChartToPdf
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a scatter chart (X values in column A, Y values in column B)
            sheet.Cells["A1"].PutValue("X");
            sheet.Cells["B1"].PutValue("Y");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue(2);
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue(4);
            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue(6);

            // Add a scatter chart to the worksheet.
            // Parameters: chart type, upper‑left row, upper‑left column, lower‑right row, lower‑right column
            int chartIdx = sheet.Charts.Add(ChartType.Scatter, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIdx];

            // Set the data source for the chart.
            // Y values (values) are taken from column B, X values (categories) from column A.
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Retrieve the actual size of the chart in pixels (returns int[2] => width, height).
            int[] actualSizeArray = chart.GetActualSize();
            Size actualSize = new Size(actualSizeArray[0], actualSizeArray[1]);

            // Convert pixel dimensions to inches using the current DPI setting.
            double dpi = CellsHelper.DPI; // DPI of the machine (default 96)
            float widthInches = (float)(actualSize.Width / dpi);
            float heightInches = (float)(actualSize.Height / dpi);

            // Export the chart to PDF, preserving its dimensions.
            // The chart will be centered on the page.
            chart.ToPdf(
                "ScatterChart.pdf",
                widthInches,
                heightInches,
                PageLayoutAlignmentType.Center,
                PageLayoutAlignmentType.Center
            );

            Console.WriteLine("Scatter chart exported to ScatterChart.pdf with original dimensions.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
