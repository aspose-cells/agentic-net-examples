using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartSaveExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["A3"].PutValue("Bananas");
            sheet.Cells["A4"].PutValue("Cherries");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(45);
            sheet.Cells["B4"].PutValue(25);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Customize the chart (example: set title)
            chart.Title.Text = "Fruit Sales";

            // Save the workbook to XLSX format (lifecycle rule: save)
            workbook.Save("CustomizedChart.xlsx", SaveFormat.Xlsx);

            // Optional: inform the user
            Console.WriteLine("Workbook with customized chart saved as CustomizedChart.xlsx");
        }
    }
}