using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsChartExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data:
            // Column A – original categories (used for the chart axis)
            // Column B – values for the series
            // Column C – custom category names to be shown as data labels
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C1"].PutValue("Label");
            sheet.Cells["C2"].PutValue("Alpha");
            sheet.Cells["C3"].PutValue("Beta");
            sheet.Cells["C4"].PutValue("Gamma");

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the series data (values) and the category axis data (original categories)
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;               // optional: show the numeric value
            series.DataLabels.ShowCellRange = true;           // use cell range as data label text
            series.DataLabels.LinkedSource = "C2:C4";         // custom labels from column C

            // Optional: format the data labels (e.g., font color)
            series.DataLabels.Font.Color = Color.Blue;

            // Save the workbook
            workbook.Save("ChartWithCustomCategoryLabels.xlsx", SaveFormat.Xlsx);
        }
    }
}