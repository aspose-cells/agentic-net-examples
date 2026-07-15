using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class AddStandardDeviationErrorBar
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data (X values in column A, Y values in column B)
        cells["A1"].PutValue("X");
        cells["B1"].PutValue("Y");
        cells["A2"].PutValue(1);
        cells["A3"].PutValue(2);
        cells["A4"].PutValue(3);
        cells["A5"].PutValue(4);
        cells["B2"].PutValue(10);
        cells["B3"].PutValue(20);
        cells["B4"].PutValue(15);
        cells["B5"].PutValue(25);

        // Add a line chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the series (Y values) and categories (X values)
        chart.NSeries.Add("B2:B5", true);
        chart.NSeries.CategoryData = "A2:A5";

        // Configure the Y‑direction error bar to display standard deviation
        Series series = chart.NSeries[0];
        series.YErrorBar.Type = ErrorBarType.StDev;               // Use standard deviation
        series.YErrorBar.DisplayType = ErrorBarDisplayType.Both; // Show both plus and minus bars

        // Save the workbook to an XLSX file
        workbook.Save("LineChart_With_StdDevErrorBar.xlsx", SaveFormat.Xlsx);
    }
}