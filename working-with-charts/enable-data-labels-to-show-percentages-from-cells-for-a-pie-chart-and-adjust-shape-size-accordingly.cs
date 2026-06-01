using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsPieChartDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pie chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["A4"].PutValue("Banana");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(50);
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["B4"].PutValue(20);

            // Add a pie chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the DataLabels of the first series
            DataLabels dataLabels = chart.NSeries[0].DataLabels;

            // Show percentage values (calculated from the data) and hide raw values
            dataLabels.ShowPercentage = true;
            dataLabels.ShowValue = false;

            // Prevent the label shape from auto‑resizing to fit the text
            dataLabels.IsResizeShapeToFitText = false;

            // Manually set a size that comfortably fits the percentage text
            dataLabels.WidthPixel = 80;   // width in pixels
            dataLabels.HeightPixel = 30;  // height in pixels

            // Save the workbook to a file
            workbook.Save("PieChartWithPercentageLabels.xlsx");
        }
    }
}