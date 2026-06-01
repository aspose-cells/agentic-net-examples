using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartOnHiddenSheet
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a new worksheet that will be hidden initially
            int hiddenSheetIndex = workbook.Worksheets.Add();
            Worksheet hiddenSheet = workbook.Worksheets[hiddenSheetIndex];
            hiddenSheet.Name = "HiddenData";

            // Hide the worksheet
            hiddenSheet.IsVisible = false; // or hiddenSheet.VisibilityType = VisibilityType.Hidden;

            // Populate sample data (categories, values, and label texts)
            hiddenSheet.Cells["A1"].PutValue("Category");
            hiddenSheet.Cells["B1"].PutValue("Value");
            hiddenSheet.Cells["C1"].PutValue("Label");
            hiddenSheet.Cells["A2"].PutValue("A");
            hiddenSheet.Cells["A3"].PutValue("B");
            hiddenSheet.Cells["A4"].PutValue("C");
            hiddenSheet.Cells["B2"].PutValue(10);
            hiddenSheet.Cells["B3"].PutValue(20);
            hiddenSheet.Cells["B4"].PutValue(30);
            hiddenSheet.Cells["C2"].PutValue("First");
            hiddenSheet.Cells["C3"].PutValue("Second");
            hiddenSheet.Cells["C4"].PutValue("Third");

            // Add a column chart to the hidden worksheet
            int chartIndex = hiddenSheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = hiddenSheet.Charts[chartIndex];

            // Set the data range for the chart (values)
            chart.NSeries.Add("B2:B4", true);
            // Set category (X‑axis) data
            chart.NSeries.CategoryData = "A2:A4";

            // Enable cell‑based data labels
            var series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;          // show numeric values
            series.DataLabels.ShowCellRange = true;      // use cell range for labels
            series.DataLabels.LinkedSource = "C2:C4";    // link to label texts

            // Make the worksheet visible again
            hiddenSheet.IsVisible = true; // or hiddenSheet.VisibilityType = VisibilityType.Visible;

            // Save the workbook
            workbook.Save("ChartOnHiddenWorksheet.xlsx", SaveFormat.Xlsx);
        }
    }
}