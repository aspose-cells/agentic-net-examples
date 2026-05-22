using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartPdfCenter
{
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

            // Define page size (in inches) and center alignment for both axes
            float pageWidth = 8.5f;   // Standard Letter width
            float pageHeight = 11f;   // Standard Letter height

            // Export the chart to PDF with centered alignment
            chart.ToPdf("CenteredChart.pdf", pageWidth, pageHeight,
                        PageLayoutAlignmentType.Center, PageLayoutAlignmentType.Center);

            Console.WriteLine("Chart has been exported to PDF with centered layout.");
        }
    }
}