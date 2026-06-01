using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AxisTitleExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";      // Categories

            // Configure X (Category) axis title
            chart.CategoryAxis.Title.Text = "Categories";
            chart.CategoryAxis.Title.IsVisible = true;
            chart.CategoryAxis.Title.Font.IsBold = true;
            chart.CategoryAxis.Title.Font.Size = 12;

            // Configure Y (Value) axis title
            chart.ValueAxis.Title.Text = "Values";
            chart.ValueAxis.Title.IsVisible = true;
            chart.ValueAxis.Title.Font.IsBold = true;
            chart.ValueAxis.Title.Font.Size = 12;

            // Save the workbook to a file
            workbook.Save("AxisTitles_Output.xlsx");
        }
    }
}