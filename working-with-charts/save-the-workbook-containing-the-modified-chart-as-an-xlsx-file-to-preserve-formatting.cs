using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartSaveExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Apples");
            worksheet.Cells["A3"].PutValue("Bananas");
            worksheet.Cells["A4"].PutValue("Cherries");

            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(30);
            worksheet.Cells["B3"].PutValue(45);
            worksheet.Cells["B4"].PutValue(25);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 6, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Modify chart properties (e.g., title)
            chart.Title.Text = "Fruit Sales";

            // Save the workbook as XLSX to preserve formatting
            workbook.Save("ModifiedChart.xlsx", SaveFormat.Xlsx);
        }
    }
}