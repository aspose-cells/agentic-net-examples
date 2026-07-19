// Title: Export Scatter Chart to PDF with Exact Size Using Aspose.Cells for .NET
// Description: Creates a workbook, fills X‑Y data, adds a Scatter chart, reads its pixel size via GetActualSize, converts pixels to inches using the system DPI, and calls Chart.ToPdf with the calculated width, height and centered alignment to generate a PDF that matches the chart's original dimensions and resolution.
// Keywords: Aspose.Cells | C# | Scatter chart PDF export | Chart.ToPdf | preserve chart size | convert pixels to inches | DPI handling | Excel to PDF conversion | .NET chart rendering
// Common Searches: export scatter chart to PDF Aspose.Cells | keep original chart dimensions when converting Excel to PDF | how to use GetActualSize with Chart.ToPdf | set DPI for chart PDF export in C# | center chart on PDF page Aspose.Cells
// Developer Intent: Generate a PDF file of a scatter chart that retains the chart's original pixel dimensions and resolution.
// Use Cases: Produce printable reports where charts must appear at exact size for branding consistency. | Automate dashboard generation that requires high‑fidelity chart images in PDF format. | Batch export charts from multiple workbooks while preserving visual quality and alignment.
// AI Prompts: Write C# code with Aspose.Cells to export any Excel chart to PDF while maintaining its size and DPI. | Explain the steps to convert a chart's pixel dimensions to inches for PDF export using Aspose.Cells. | Show how to loop through all charts on a worksheet and save each as a centered PDF with original dimensions.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace ScatterChartPdfExport
{
    // Creates a workbook, fills X‑Y data, adds a Scatter chart, reads its pixel size via GetActualSize, converts pixels to inches using the system DPI, and calls Chart.ToPdf with the calculated width, height and centered alignment to generate a PDF that matches the chart's original dimensions and resolution.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for a scatter chart (X values in column A, Y values in column B)
                worksheet.Cells["A1"].PutValue("X");
                worksheet.Cells["B1"].PutValue("Y");
                worksheet.Cells["A2"].PutValue(1);
                worksheet.Cells["B2"].PutValue(2);
                worksheet.Cells["A3"].PutValue(2);
                worksheet.Cells["B3"].PutValue(4);
                worksheet.Cells["A4"].PutValue(3);
                worksheet.Cells["B4"].PutValue(6);
                worksheet.Cells["A5"].PutValue(4);
                worksheet.Cells["B5"].PutValue(8);

                // Add a scatter chart to the worksheet
                // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column
                int chartIndex = worksheet.Charts.Add(ChartType.Scatter, 5, 0, 20, 15);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B5", true);          // Y values
                chart.NSeries.CategoryData = "A2:A5";      // X values

                // Optional: give the chart a title
                chart.Title.Text = "Sample Scatter Chart";

                // Retrieve the actual size of the chart in pixels (returns int[2] => width, height)
                int[] actualSize = chart.GetActualSize();

                // Determine the DPI of the current environment (default is 96)
                float dpi = (float)CellsHelper.DPI;

                // Convert pixel dimensions to inches to preserve the exact size in the PDF
                float widthInInches = actualSize[0] / dpi;
                float heightInInches = actualSize[1] / dpi;

                // Export the chart to a PDF file, preserving its dimensions and using centered alignment
                chart.ToPdf(
                    "ScatterChart.pdf",
                    widthInInches,
                    heightInInches,
                    PageLayoutAlignmentType.Center,
                    PageLayoutAlignmentType.Center
                );

                Console.WriteLine("Scatter chart exported to PDF with original dimensions.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
