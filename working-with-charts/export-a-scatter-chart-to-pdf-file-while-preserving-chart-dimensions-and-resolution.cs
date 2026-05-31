using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ScatterChartToPdfDemo
{
    class Program
    {
        static void Main()
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
            worksheet.Cells["A6"].PutValue(5);
            worksheet.Cells["B6"].PutValue(10);

            // Add a scatter chart to the worksheet
            // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column
            int chartIndex = worksheet.Charts.Add(ChartType.Scatter, 5, 0, 20, 15);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B6", true);          // Y values
            chart.NSeries.CategoryData = "A2:A6";      // X values

            // Optional: set chart title and style
            chart.Title.Text = "Sample Scatter Chart";
            chart.Style = 2; // Built‑in style

            // Export the chart to PDF.
            // The overload allows us to specify the page size (in inches) and alignment,
            // which helps preserve the chart's visual dimensions.
            float pageWidthInches = 8.5f;   // Standard letter width
            float pageHeightInches = 11f;  // Standard letter height
            chart.ToPdf(
                "ScatterChart.pdf",
                pageWidthInches,
                pageHeightInches,
                PageLayoutAlignmentType.Center,
                PageLayoutAlignmentType.Center);

            Console.WriteLine("Scatter chart exported to ScatterChart.pdf successfully.");
        }
    }
}