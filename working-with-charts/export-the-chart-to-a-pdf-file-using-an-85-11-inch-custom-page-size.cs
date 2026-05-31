using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartToPdf
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Fruits");
            sheet.Cells["A3"].PutValue("Vegetables");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(50);
            sheet.Cells["B3"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Export the chart to a PDF with a custom 8.5 x 11 inch page size,
            // centering the chart both horizontally and vertically.
            chart.ToPdf(
                "ChartCustomSize.pdf",          // output file name
                8.5f,                          // desired page width in inches
                11f,                           // desired page height in inches
                PageLayoutAlignmentType.Center, // horizontal alignment
                PageLayoutAlignmentType.Center  // vertical alignment
            );

            Console.WriteLine("Chart exported to PDF with custom page size successfully.");
        }
    }
}