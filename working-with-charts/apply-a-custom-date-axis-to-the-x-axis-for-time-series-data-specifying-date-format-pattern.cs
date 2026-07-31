// Title: C# – Apply Custom Date Format to X Axis of Time‑Scale Chart using Aspose.Cells
// Description: Creates a workbook, adds date/value rows, inserts a line chart, sets the category axis to TimeScale with a monthly base unit, and applies a custom date pattern (e.g., "MMM dd, yyyy") to the X‑axis via XValuesFormatCode or TickLabels.NumberFormat before saving.
// Keywords: Aspose.Cells | C# | .NET | chart | line chart | time scale axis | custom date format | X axis date pattern | CategoryAxis | XValuesFormatCode | TickLabels.NumberFormat | date axis formatting
// Common Searches: Aspose.Cells set X axis date format | C# time scale chart custom date pattern | How to format chart axis dates in Aspose.Cells | Apply custom date format to category axis Aspose.Cells .NET
// Developer Intent: The developer wants to display dates on a chart’s X‑axis using a specific format while using a time‑scale axis.
// Use Cases: Monthly sales line chart showing each month as "Jan 01, 2023" on the X axis. | Project timeline chart with milestone dates formatted as "MMM dd, yyyy" for clear reporting. | Financial performance chart where quarterly dates are displayed in a custom pattern.
// AI Prompts: Provide C# code that sets a custom date format on the category axis of a time‑scale line chart using Aspose.Cells. | Show how to change the base unit of a time‑scale axis to weeks and apply a "yyyy-MM-dd" label format in Aspose.Cells. | Explain the difference between using XValuesFormatCode and TickLabels.NumberFormat for date label formatting on a chart axis.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds date/value rows, inserts a line chart, sets the category axis to TimeScale with a monthly base unit, and applies a custom date pattern (e.g., "MMM dd, yyyy") to the X‑axis via XValuesFormatCode or TickLabels.NumberFormat before saving.
class CustomDateAxisDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample date and value data
        sheet.Cells["A1"].PutValue("Date");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["A3"].PutValue(new DateTime(2023, 2, 1));
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["A4"].PutValue(new DateTime(2023, 3, 1));
        sheet.Cells["B4"].PutValue(200);

        // Add a line chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Set the chart data source
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Dates for X axis

        // Configure the X (category) axis as a time scale
        chart.CategoryAxis.CategoryType = CategoryType.TimeScale;
        chart.CategoryAxis.BaseUnitScale = TimeUnit.Months; // optional base unit

        // Apply a custom date format to the X axis
        // Using the series property
        chart.NSeries[0].XValuesFormatCode = "mmm dd, yyyy";

        // Or directly via tick label format (both achieve the same result)
        chart.CategoryAxis.TickLabels.NumberFormat = "mmm dd, yyyy";

        // Save the workbook
        workbook.Save("CustomDateAxisDemo.xlsx");
    }
}
