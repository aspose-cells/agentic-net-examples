using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsShowCellRangeDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["C2"].PutValue("100 units");
            sheet.Cells["C3"].PutValue("200 units");

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Define the series data range and category data
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Access the first series
            Series series = chart.NSeries[0];

            // Enable data labels and activate cell‑based data labels
            series.DataLabels.ShowValue = true;          // optional: show the numeric value
            series.DataLabels.ShowCellRange = true;      // activate cell range as data labels
            series.DataLabels.LinkedSource = "C2:C3";    // link to cells containing custom label text
            series.DataLabels.Font.Color = Color.Blue;  // optional styling

            // Save the workbook
            workbook.Save("ShowCellRangeDemo.xlsx");
        }
    }
}