// Title: How to convert a line chart’s X‑axis from category to date (time‑scale) axis using Aspose.Cells for .NET (C#)
// AI Prompts: Create a line chart with Aspose.Cells in C# and set the X‑axis to a time‑scale axis with monthly major units and weekly minor units. | Modify an existing Aspose.Cells chart to use CategoryType.TimeScale and apply the label format "mmm yyyy" for the date axis. | Generate an Excel workbook that contains a date‑axis line chart, configuring base unit, major/minor units, and custom X‑axis label formatting via the Aspose.Cells API.
// Common Searches: Aspose.Cells C# line chart date axis monthly major unit weekly minor unit | Set CategoryType to TimeScale in Aspose.Cells chart example | Format X axis labels as month year in Aspose.Cells line chart C#
// Tags: Aspose.Cells line chart time‑scale axis configuration | C# set chart CategoryType TimeScale Aspose.Cells | date axis major unit months Aspose.Cells | chart minor unit weeks Aspose.Cells | X axis label format mmm yyyy Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDateAxisDemo
{
    // The example creates a workbook, fills column A with dates and column B with numeric values, adds a line chart, assigns the data ranges, switches the X‑axis to a time‑scale axis with monthly major units and weekly minor units, formats the axis labels as "mmm yyyy", and saves the file as DateAxisChart.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with dates in column A and numeric values in column B
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
            int chartIndex = sheet.Charts.Add(ChartType.Line, 6, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series (values) and the category (dates)
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Switch the X axis from a generic category axis to a date (time) axis
            chart.CategoryAxis.CategoryType = CategoryType.TimeScale;   // Enable time scaling
            chart.CategoryAxis.BaseUnitScale = TimeUnit.Months;        // Base unit (optional)
            chart.CategoryAxis.MajorUnitScale = TimeUnit.Months;       // Major grid every month
            chart.CategoryAxis.MajorUnit = 1;                         // One month per major unit
            chart.CategoryAxis.MinorUnitScale = TimeUnit.Days;        // Minor grid daily
            chart.CategoryAxis.MinorUnit = 7;                         // One week per minor unit

            // Optional: format the X‑axis labels to show month/year
            chart.NSeries[0].XValuesFormatCode = "mmm yyyy";

            // Save the workbook to an XLSX file
            workbook.Save("DateAxisChart.xlsx");
        }
    }
}
