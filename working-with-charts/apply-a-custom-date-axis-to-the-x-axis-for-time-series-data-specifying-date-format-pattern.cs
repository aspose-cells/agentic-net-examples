// Title: Custom Date Axis Formatting for Time‑Scale Line Charts with Aspose.Cells .NET
// Description: Shows how to create an Excel workbook, fill it with monthly dates and numeric values, add a line chart, switch the category axis to a time scale, and apply a custom date pattern (e.g., dd‑MMM‑yyyy) to the X‑axis tick labels via the XValuesFormatCode property or the axis TickLabels.NumberFormat setting.
// Keywords: Aspose.Cells | C# | custom date axis | time scale chart | category axis format | XValuesFormatCode | TickLabels.NumberFormat | line chart formatting | Excel chart date pattern | date axis .NET
// Common Searches: Aspose.Cells set custom date format on chart X axis | time scale category axis date pattern Aspose.Cells | format X axis dates in line chart C# | apply dd-MMM-yyyy to chart axis Aspose.Cells | how to use XValuesFormatCode for date axis
// Developer Intent: The developer wants the X‑axis of a time‑scale chart to display dates using a specific format.
// Use Cases: Monthly sales line chart where the X‑axis shows dates as dd‑MMM‑yyyy. | Financial time‑series report with a line chart that uses a time‑scale axis and custom date labels. | Automated Excel export that requires precise date formatting on chart category axes.
// AI Prompts: Generate C# code with Aspose.Cells to create a line chart that uses a time‑scale X axis and formats tick labels as yyyy/MM/dd. | Explain how to set a custom date pattern for a chart’s category axis in Aspose.Cells without using XValuesFormatCode. | Provide an Aspose.Cells example that formats the X axis of a time‑scale chart to display only month names.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to create an Excel workbook, fill it with monthly dates and numeric values, add a line chart, switch the category axis to a time scale, and apply a custom date pattern (e.g., dd‑MMM‑yyyy) to the X‑axis tick labels via the XValuesFormatCode property or the axis TickLabels.NumberFormat setting.
class CustomDateAxisDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add headers for date and value columns
        sheet.Cells["A1"].PutValue("Date");
        sheet.Cells["B1"].PutValue("Value");

        // Populate sample date (monthly) and numeric data
        DateTime startDate = new DateTime(2023, 1, 1);
        for (int i = 0; i < 5; i++)
        {
            sheet.Cells[i + 2, 0].PutValue(startDate.AddMonths(i)); // Column A: dates
            sheet.Cells[i + 2, 1].PutValue((i + 1) * 10);          // Column B: values
        }

        // Add a line chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Set the chart data range: values in B2:B6, categories (X axis) in A2:A6
        chart.NSeries.Add("B2:B6", true);
        chart.NSeries.CategoryData = "A2:A6";

        // Configure the X (category) axis to use a time scale
        chart.CategoryAxis.CategoryType = CategoryType.TimeScale;

        // Apply a custom date format to the X axis tick labels
        // Using the series property (XValuesFormatCode) as per Aspose.Cells API
        chart.NSeries[0].XValuesFormatCode = "dd-MMM-yyyy";

        // Alternatively, you can set the same format via the axis tick labels
        // chart.CategoryAxis.TickLabels.NumberFormat = "dd-MMM-yyyy";

        // Save the workbook to a file
        workbook.Save("CustomDateAxisDemo.xlsx");
    }
}
