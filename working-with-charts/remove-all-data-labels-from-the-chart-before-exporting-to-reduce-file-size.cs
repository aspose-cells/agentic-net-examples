using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDataLabelRemoval
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
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

            // Enable data labels initially (optional, just to demonstrate removal)
            chart.NSeries[0].DataLabels.ShowValue = true;

            // ------------------------------------------------------------
            // Remove all data labels from the chart to reduce file size
            // ------------------------------------------------------------
            foreach (Series series in chart.NSeries)
            {
                // Mark the DataLabels object as deleted; this removes all labels for the series
                series.DataLabels.IsDeleted = true;
            }

            // Save the workbook with the chart (data labels removed)
            workbook.Save("ChartWithoutDataLabels.xlsx", SaveFormat.Xlsx);
        }
    }
}