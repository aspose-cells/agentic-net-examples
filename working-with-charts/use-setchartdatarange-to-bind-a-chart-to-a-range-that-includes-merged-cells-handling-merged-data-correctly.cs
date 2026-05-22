using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsMergedChartDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate data for the chart
            // Merge the header cells (A1:B1) to simulate a merged title row
            cells["A1"].PutValue("Sales Report");
            cells.Merge(0, 0, 1, 2); // Merge A1:B1 (row 0, column 0, 1 row, 2 columns)

            // Category labels
            cells["A2"].PutValue("Q1");
            cells["A3"].PutValue("Q2");
            cells["A4"].PutValue("Q3");
            cells["A5"].PutValue("Q4");

            // Corresponding values
            cells["B2"].PutValue(15000);
            cells["B3"].PutValue(20000);
            cells["B4"].PutValue(18000);
            cells["B5"].PutValue(22000);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Bind the chart to the range that includes the merged header.
            // SetChartDataRange will treat the merged header as a single cell (top‑left cell A1).
            chart.SetChartDataRange("A1:B5", true);

            // Ensure the chart uses the merged header as the series name.
            // The first row (A1:B1) contains the series name; the rest are categories and values.
            chart.NSeries[0].Name = "=Sheet1!$B$1";

            // Optionally, plot all cells (including hidden/merged) – default is true.
            chart.PlotVisibleCellsOnly = false;

            // Save the workbook
            workbook.Save("MergedChartDemo.xlsx");
        }
    }
}