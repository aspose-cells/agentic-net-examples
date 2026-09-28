// Title: Set a custom date format on the X‑axis of a line chart using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a line chart in an Aspose.Cells workbook, configures the category axis as a time scale, and applies a custom date format (e.g., yyyy‑MM‑dd) to the X‑axis values. | Show how to use the XValuesFormatCode property to define the display pattern of dates on a chart’s category axis in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# line chart custom X axis date format yyyy-MM-dd | How to use time scale category axis with date values in Aspose.Cells | Set XValuesFormatCode for chart axis in Aspose.Cells .NET example
// Tags: custom date format on chart axis Aspose.Cells | time scale category axis line chart .NET | XValuesFormatCode property Aspose.Cells | save workbook with formatted date axis Excel

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDateAxisDemo
{
    // Demonstrates creating a workbook, inserting date and numeric data, adding a line chart, setting the category axis to a time scale, applying a custom date format (yyyy‑MM‑dd) to the X‑axis via XValuesFormatCode, and saving the file as an Excel workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate worksheet with date (X) and numeric (Y) data
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue(new DateTime(2024, 1, 1));
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue(new DateTime(2024, 2, 1));
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["A4"].PutValue(new DateTime(2024, 3, 1));
            sheet.Cells["B4"].PutValue(180);

            // Add a line chart
            int chartIdx = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIdx];

            // Set the data range for the series (Y values) and category (X values)
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure the X axis as a time scale (optional but recommended for date axes)
            chart.CategoryAxis.CategoryType = CategoryType.TimeScale;

            // Apply a custom date format to the X values of the series
            // This controls how dates appear on the axis
            chart.NSeries[0].XValuesFormatCode = "yyyy-mm-dd";

            // Save the workbook
            workbook.Save("DateAxisCustomFormat.xlsx");
        }
    }
}
