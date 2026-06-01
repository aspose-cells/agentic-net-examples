using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsErrorBarExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the line chart
            // Column A – X axis (categories)
            cells["A1"].PutValue("Month");
            cells["A2"].PutValue("Jan");
            cells["A3"].PutValue("Feb");
            cells["A4"].PutValue("Mar");
            cells["A5"].PutValue("Apr");

            // Column B – Y values for the series
            cells["B1"].PutValue("Sales");
            cells["B2"].PutValue(120);
            cells["B3"].PutValue(150);
            cells["B4"].PutValue(180);
            cells["B5"].PutValue(210);

            // Add a line chart to the worksheet
            // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column
            int chartIndex = sheet.Charts.Add(ChartType.Line, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series (Y values) and categories (X axis)
            chart.NSeries.Add("B2:B5", true);          // Y values
            chart.NSeries.CategoryData = "A2:A5";      // X categories

            // Access the first (and only) series
            Series series = chart.NSeries[0];

            // Configure the Y-direction error bar to display standard deviation
            series.YErrorBar.Type = ErrorBarType.StDev;                 // Use standard deviation as the error amount
            series.YErrorBar.DisplayType = ErrorBarDisplayType.Both;   // Show both plus and minus error bars
            series.YErrorBar.IsVisible = true;                         // Ensure the error bars are visible

            // Optional: customize appearance of the error bars (color, line style, etc.)
            series.YErrorBar.Color = Color.DarkGray;
            series.YErrorBar.Weight = Aspose.Cells.Drawing.WeightType.SingleLine;
            series.YErrorBar.DashType = Aspose.Cells.Drawing.MsoLineDashStyle.Solid;

            // Save the workbook to an XLSX file
            workbook.Save("LineChart_With_StdDev_ErrorBars.xlsx", SaveFormat.Xlsx);
        }
    }
}