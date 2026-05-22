using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartAlignmentDemo
{
    class Program
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

            // Set the data range for the chart
            chart.SetChartDataRange("A1:B4", true);
            chart.Title.Text = "Sample Chart";

            // Export the chart to PDF, aligning it to the top‑left corner of the page
            // Page size: 8.5 x 11 inches (standard Letter)
            // Horizontal alignment: Left, Vertical alignment: Top
            chart.ToPdf("ChartTopLeft.pdf", 8.5f, 11f,
                        PageLayoutAlignmentType.Left,
                        PageLayoutAlignmentType.Top);

            Console.WriteLine("Chart exported to PDF with top‑left alignment.");
        }
    }
}