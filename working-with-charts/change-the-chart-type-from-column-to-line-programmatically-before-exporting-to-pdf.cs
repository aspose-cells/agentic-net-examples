using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartTypeChange
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
            sheet.Cells["A4"].PutValue("Grains");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(50);
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["B4"].PutValue(20);

            // Add a column chart (initial type is Column)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Change the chart type from Column to Line
            chart.Type = ChartType.Line;

            // Export the chart to a PDF file
            chart.ToPdf("ChartLine.pdf");

            Console.WriteLine("Chart type changed to Line and exported to PDF successfully.");
        }
    }
}