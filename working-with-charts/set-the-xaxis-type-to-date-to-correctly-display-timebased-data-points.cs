using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SetXAxisDateScale
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample date and value data
        worksheet.Cells["A1"].PutValue("Date");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue(new DateTime(2024, 1, 1));
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["A3"].PutValue(new DateTime(2024, 2, 1));
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["A4"].PutValue(new DateTime(2024, 3, 1));
        worksheet.Cells["B4"].PutValue(30);

        // Add a line chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Line, 5, 0, 25, 10);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data source for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Configure the X‑axis (category axis) to use a date (time) scale
        chart.CategoryAxis.CategoryType = CategoryType.TimeScale;

        // Optional: define the base unit and major/minor units for better display
        chart.CategoryAxis.BaseUnitScale = TimeUnit.Months;
        chart.CategoryAxis.MajorUnitScale = TimeUnit.Months;
        chart.CategoryAxis.MajorUnit = 1;      // one month per major tick
        chart.CategoryAxis.MinorUnitScale = TimeUnit.Days;
        chart.CategoryAxis.MinorUnit = 7;      // one week per minor tick

        // Save the workbook
        workbook.Save("ChartWithDateAxis.xlsx");
    }
}