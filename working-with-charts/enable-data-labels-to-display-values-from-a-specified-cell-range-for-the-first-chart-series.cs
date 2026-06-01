using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsDataLabelsFromRange
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

            // Cells that contain custom label text (the range to be shown as data labels)
            sheet.Cells["C2"].PutValue("10 units");
            sheet.Cells["C3"].PutValue("20 units");
            sheet.Cells["C4"].PutValue("30 units");

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series (values) and categories
            chart.NSeries.Add("=Sheet1!$B$2:$B$4", true);
            chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$4";

            // Access the first series
            Series firstSeries = chart.NSeries[0];

            // Enable data labels and configure them to display values from the specified cell range
            firstSeries.DataLabels.ShowValue = true;          // Show the numeric value (optional)
            firstSeries.DataLabels.ShowCellRange = true;     // Enable using a cell range for labels
            firstSeries.DataLabels.LinkedSource = "C2:C4";   // Range that contains the label texts
            firstSeries.DataLabels.Font.Color = Color.Blue; // Example styling

            // Save the workbook
            workbook.Save("ChartWithDataLabelsFromRange.xlsx", SaveFormat.Xlsx);
        }
    }
}