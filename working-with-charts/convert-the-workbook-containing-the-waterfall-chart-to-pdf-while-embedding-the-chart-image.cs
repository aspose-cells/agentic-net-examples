using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsWaterfallToPdf
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the Waterfall chart
            // Category column
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Start");
            sheet.Cells["A3"].PutValue("Revenue");
            sheet.Cells["A4"].PutValue("Cost");
            sheet.Cells["A5"].PutValue("Profit");

            // Values column
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(5000);
            sheet.Cells["B3"].PutValue(8000);
            sheet.Cells["B4"].PutValue(-3000);
            sheet.Cells["B5"].PutValue(0); // Final total will be calculated automatically

            // Add a Waterfall chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart (including both categories and values)
            chart.SetChartDataRange("A1:B5", true);
            chart.Title.Text = "Waterfall Chart Example";

            // Export the chart (which includes the Waterfall image) to a PDF file
            // This uses the Chart.ToPdf(string) method as defined in the documentation
            chart.ToPdf("WaterfallChart.pdf");

            Console.WriteLine("Waterfall chart has been exported to PDF successfully.");
        }
    }
}