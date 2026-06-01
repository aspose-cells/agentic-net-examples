using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class DateAxisChartExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate worksheet with date‑based categories and numeric values
        sheet.Cells["A1"].PutValue("Date");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue(new DateTime(2024, 1, 1));
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue(new DateTime(2024, 2, 1));
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue(new DateTime(2024, 3, 1));
        sheet.Cells["B4"].PutValue(30);
        sheet.Cells["A5"].PutValue(new DateTime(2024, 4, 1));
        sheet.Cells["B5"].PutValue(40);

        // Add a line chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Line, 6, 0, 20, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the series and the category (X) axis
        chart.NSeries.Add("B2:B5", true);
        chart.NSeries.CategoryData = "A2:A5";

        // Change the category axis to a date (time) axis
        chart.CategoryAxis.CategoryType = CategoryType.TimeScale;

        // Optional: define the base unit scale for better tick spacing (e.g., months)
        chart.CategoryAxis.BaseUnitScale = TimeUnit.Months;
        chart.CategoryAxis.MajorUnitScale = TimeUnit.Months;
        chart.CategoryAxis.MajorUnit = 1; // one month per major tick

        // Save the workbook to an XLSX file
        workbook.Save("DateAxisChart.xlsx", SaveFormat.Xlsx);
    }
}