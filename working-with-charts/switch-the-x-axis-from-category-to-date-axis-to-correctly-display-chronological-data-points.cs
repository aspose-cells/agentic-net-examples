using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SwitchXAxisToDateAxis
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate worksheet with chronological data
        sheet.Cells["A1"].PutValue("Date");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue(new DateTime(2024, 1, 1));
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue(new DateTime(2024, 2, 1));
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue(new DateTime(2024, 3, 1));
        sheet.Cells["B4"].PutValue(30);

        // Add a line chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 25, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data source for the chart
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";     // Dates (X‑axis)

        // Switch the X axis from a categorical axis to a date (time) axis
        chart.CategoryAxis.CategoryType = CategoryType.TimeScale;

        // Optional: configure the time scale (e.g., months)
        chart.CategoryAxis.BaseUnitScale = TimeUnit.Months;
        chart.CategoryAxis.MajorUnitScale = TimeUnit.Months;
        chart.CategoryAxis.MajorUnit = 1;

        // Save the workbook with the configured chart
        workbook.Save("DateAxisChart.xlsx");
    }
}