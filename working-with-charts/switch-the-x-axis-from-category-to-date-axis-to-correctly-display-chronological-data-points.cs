// Title: Aspose.Cells C# – Switch Chart X‑Axis to Date (Time‑Scale) Axis
// Description: This C# example shows how to create a workbook with date values, add a line chart, and convert the X‑axis from a categorical axis to a time‑scale (date) axis using Aspose.Cells. It also demonstrates configuring base, major and minor units (months and days) before saving the file as an Excel workbook.
// Keywords: Aspose.Cells chart date axis C# | switch X axis to time scale Aspose.Cells | C# line chart date axis example | CategoryAxis CategoryType TimeScale | .NET chart axis time unit configuration | Aspose.Cells set base unit months | Aspose.Cells major minor tick settings | Excel date axis Aspose.Cells | GitHub Aspose.Cells chart sample | US developers Aspose.Cells chart tutorial
// Common Searches: How to change a chart X axis to a date axis in Aspose.Cells C# | Aspose.Cells time‑scale axis configuration .NET | Set major and minor units on a date axis using Aspose.Cells | C# example for date‑scaled line chart with Aspose.Cells | Aspose.Cells chart axis base unit months
// Developer Intent: Convert a chart’s X‑axis to a date (time‑scale) axis and adjust its time‑unit settings in Aspose.Cells for .NET.
// Use Cases: Monthly sales trend line chart with chronological dates on the X‑axis. | Financial performance report where the axis shows month‑level major ticks and weekly minor ticks. | Automated generation of Excel workbooks that require accurate timeline visualization for project milestones.
// AI Prompts: Generate C# code that changes a chart’s category axis to a time‑scale axis and sets major ticks every month using Aspose.Cells. | Explain how to configure base, major, and minor unit scales for a date axis in an Aspose.Cells line chart. | Provide a step‑by‑step guide to add multiple series to a date‑axis chart and format the axis labels for month and day display.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This C# example shows how to create a workbook with date values, add a line chart, and convert the X‑axis from a categorical axis to a time‑scale (date) axis using Aspose.Cells. It also demonstrates configuring base, major and minor units (months and days) before saving the file as an Excel workbook.
class SwitchXAxisToDateAxis
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
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue(new DateTime(2024, 2, 1));
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue(new DateTime(2024, 3, 1));
        sheet.Cells["B4"].PutValue(30);
        sheet.Cells["A5"].PutValue(new DateTime(2024, 4, 1));
        sheet.Cells["B5"].PutValue(40);

        // Add a line chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 25, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the series and the category (X) axis
        chart.NSeries.Add("B2:B5", true);
        chart.NSeries.CategoryData = "A2:A5";

        // Switch the X axis from a categorical axis to a date (time) axis
        chart.CategoryAxis.CategoryType = CategoryType.TimeScale;

        // Optional: configure the time scale units for better display
        chart.CategoryAxis.BaseUnitScale = TimeUnit.Months;   // Base unit for the axis
        chart.CategoryAxis.MajorUnitScale = TimeUnit.Months; // Major tick every month
        chart.CategoryAxis.MajorUnit = 1;
        chart.CategoryAxis.MinorUnitScale = TimeUnit.Days;   // Minor tick every week
        chart.CategoryAxis.MinorUnit = 7;

        // Save the workbook with the configured chart
        workbook.Save("ChartWithDateAxis.xlsx");
    }
}
