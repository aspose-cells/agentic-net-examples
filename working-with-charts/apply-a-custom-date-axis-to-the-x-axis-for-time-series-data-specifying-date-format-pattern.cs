// Title: Aspose.Cells .NET – Set Custom Date Format on Chart X‑Axis (Time Scale)
// Description: Creates a workbook, fills column A with weekly dates starting 1‑Jan‑2023 and column B with numeric values, adds a line chart, configures the X‑axis as a time‑scale category axis, applies the date pattern "dd-MMM-yyyy" to the axis labels, and saves the file as CustomDateAxisDemo.xlsx.
// Keywords: Aspose.Cells custom date axis | C# chart X axis date format | time scale category axis Aspose.Cells | line chart date labels .NET | Aspose.Cells chart formatting
// Common Searches: Aspose.Cells set custom date format on chart X axis | How to use TimeScale category axis in Aspose.Cells | Change date label pattern in Aspose.Cells chart | C# Aspose.Cells chart base unit days
// Developer Intent: The developer wants to generate a line chart whose X‑axis uses a time‑scale and displays dates in a specific format.
// Use Cases: Weekly sales trend chart with dates shown as dd‑MMM‑yyyy for clear reporting. | Project timeline visualization where milestones are plotted on a formatted date axis. | Financial performance chart that requires consistent date patterns across the X‑axis.
// AI Prompts: Show C# code to apply a custom date format to the X‑axis of an Aspose.Cells chart. | How can I set the CategoryAxis to TimeScale and change its base unit to months with format MMM yyyy? | Explain steps to configure a line chart’s X‑axis as a time scale and format date labels in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, fills column A with weekly dates starting 1‑Jan‑2023 and column B with numeric values, adds a line chart, configures the X‑axis as a time‑scale category axis, applies the date pattern "dd-MMM-yyyy" to the axis labels, and saves the file as CustomDateAxisDemo.xlsx.
class CustomDateAxisDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add headers
        sheet.Cells["A1"].PutValue("Date");
        sheet.Cells["B1"].PutValue("Value");

        // Populate sample date and value data
        DateTime startDate = new DateTime(2023, 1, 1);
        for (int i = 0; i < 10; i++)
        {
            // Dates spaced one week apart
            sheet.Cells[i + 2, 0].PutValue(startDate.AddDays(i * 7));
            // Sample numeric values
            sheet.Cells[i + 2, 1].PutValue((i + 1) * 10);
        }

        // Add a line chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Set the chart data range
        chart.NSeries.Add("B2:B11", true);          // Values
        chart.NSeries.CategoryData = "A2:A11";      // Dates (X axis)

        // Configure the X (category) axis as a time scale
        chart.CategoryAxis.CategoryType = CategoryType.TimeScale;
        chart.CategoryAxis.BaseUnitScale = TimeUnit.Days; // optional base unit

        // Apply a custom date format to the X values
        chart.NSeries[0].XValuesFormatCode = "dd-MMM-yyyy";

        // Save the workbook
        workbook.Save("CustomDateAxisDemo.xlsx");
    }
}
