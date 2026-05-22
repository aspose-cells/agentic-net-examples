using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartSubtitleDemo
{
    class Program
    {
        static void Main(string[] args)
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
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set the main title (optional)
            chart.Title.Text = "Sample Chart Title";

            // Set the subtitle to a custom string
            chart.SubTitle.Text = "Custom Chart Subtitle";

            // Verify the subtitle by reading the property
            string subtitleValue = chart.SubTitle.Text;
            Console.WriteLine("Subtitle set to: " + subtitleValue);

            // Save the workbook to a file
            workbook.Save("ChartWithCustomSubtitle.xlsx");
        }
    }
}