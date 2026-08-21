// Title: Set a DateAxis (TimeScale) for the Category Axis in an Aspose.Cells Line Chart (C#)
// Description: Creates a workbook, adds date values to column A and numeric values to column B, builds a line chart, binds the series and category ranges, then converts the X‑axis to a DateAxis by setting CategoryType to TimeScale and configuring BaseUnitScale and MajorUnit for monthly intervals before saving as XLSX.
// Keywords: Aspose.Cells DateAxis | CategoryType TimeScale C# | line chart time‑scale axis | BaseUnitScale months | Aspose.Cells chart axis configuration | C# Excel chart date axis
// Common Searches: Aspose.Cells set category axis to DateAxis .NET | How to use TimeScale axis in Aspose.Cells chart | C# line chart with monthly dates on X‑axis | Configure BaseUnitScale and MajorUnit in Aspose.Cells
// Developer Intent: Convert the chart’s category axis to a time‑based DateAxis so that date values are rendered correctly on the X‑axis.
// Use Cases: Monthly sales trend line chart with dates on the X‑axis. | Project timeline visualization where milestones align to calendar dates. | Financial performance chart that automatically groups data by months or years.
// AI Prompts: Show a complete C# example that sets CategoryAxis.CategoryType to TimeScale and defines monthly BaseUnitScale and MajorUnit in Aspose.Cells. | Explain how to bind worksheet date cells to a chart’s CategoryData and then switch the axis to a DateAxis. | Describe how to change the time unit (days, months, years) of a DateAxis in Aspose.Cells and its impact on chart rendering.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds date values to column A and numeric values to column B, builds a line chart, binds the series and category ranges, then converts the X‑axis to a DateAxis by setting CategoryType to TimeScale and configuring BaseUnitScale and MajorUnit for monthly intervals before saving as XLSX.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add date (X‑axis) and numeric (Y‑axis) data
        sheet.Cells["A1"].PutValue("Date");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue(new DateTime(2024, 1, 1));
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue(new DateTime(2024, 2, 1));
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue(new DateTime(2024, 3, 1));
        sheet.Cells["B4"].PutValue(30);

        // Add a line chart
        int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 25, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the series data and the category (X‑axis) data
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Change the category axis to a time‑based axis (DateAxis)
        chart.CategoryAxis.CategoryType = CategoryType.TimeScale;

        // Optional: configure the time unit (e.g., months) and major unit interval
        chart.CategoryAxis.BaseUnitScale = TimeUnit.Months;
        chart.CategoryAxis.MajorUnitScale = TimeUnit.Months;
        chart.CategoryAxis.MajorUnit = 1;

        // Save the workbook
        workbook.Save("TimeBasedChart.xlsx", SaveFormat.Xlsx);
    }
}
