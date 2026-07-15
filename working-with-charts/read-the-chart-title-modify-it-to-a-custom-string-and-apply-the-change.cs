using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartTitleExample
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
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set an initial title
            chart.Title.Text = "Original Chart Title";
            chart.Title.IsVisible = true;

            // Read the current title
            string currentTitle = chart.Title.Text;
            Console.WriteLine("Current Title: " + currentTitle);

            // Modify the title to a custom string
            string customTitle = "Custom Chart Title";
            chart.Title.Text = customTitle;

            // Verify the change
            Console.WriteLine("Updated Title: " + chart.Title.Text);

            // Save the workbook
            workbook.Save("ChartTitleModified.xlsx");
        }
    }
}