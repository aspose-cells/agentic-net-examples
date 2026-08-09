// Title: C# – Aspose.Cells: Line chart with time‑scale X axis and custom date format
// Description: Shows how to create a workbook, insert a line chart, switch the category axis to a time scale, and apply a date pattern (e.g., dd‑MMM‑yyyy) to both X‑values and tick labels using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# chart date axis | time scale category axis | custom date pattern | line chart Excel | XValuesFormatCode | CategoryAxis.TickLabels | date formatting | time‑series visualization | Excel export C#
// Common Searches: Aspose.Cells set X axis date format C# | time scale chart category axis Aspose.Cells | format chart dates dd-MMM-yyyy .NET | C# line chart custom date labels | apply date pattern to chart axis Aspose.Cells
// Developer Intent: Apply a specific date pattern to the X‑axis of a time‑scale chart in a .NET workbook.
// Use Cases: Produce a monthly sales trend chart where dates appear as "01‑Jan‑2023", "01‑Feb‑2023", etc. | Generate a financial performance line graph with uniformly formatted timestamps for reporting dashboards. | Export a time‑series dataset to Excel while preserving readable date labels without modifying source cells.
// AI Prompts: Give me C# code that configures a line chart’s CategoryAxis as a TimeScale and formats the X‑axis dates with "dd-MMM-yyyy" using Aspose.Cells. | How can I set a custom date pattern for both XValues and tick labels of a chart in Aspose.Cells for .NET? | Explain the steps to apply a date format to a chart axis without changing the underlying worksheet data.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to create a workbook, insert a line chart, switch the category axis to a time scale, and apply a date pattern (e.g., dd‑MMM‑yyyy) to both X‑values and tick labels using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add header
        sheet.Cells["A1"].PutValue("Date");
        sheet.Cells["B1"].PutValue("Value");

        // Populate date and value data
        sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["A3"].PutValue(new DateTime(2023, 2, 1));
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["A4"].PutValue(new DateTime(2023, 3, 1));
        sheet.Cells["B4"].PutValue(200);

        // Add a line chart
        int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Set the chart data source
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Dates for X axis

        // Configure the X (category) axis as a time scale
        chart.CategoryAxis.CategoryType = CategoryType.TimeScale;

        // Apply a custom date format to the X axis values
        chart.NSeries[0].XValuesFormatCode = "dd-MMM-yyyy";

        // Also set the tick label format for consistency
        chart.CategoryAxis.TickLabels.NumberFormat = "dd-MMM-yyyy";

        // Save the workbook
        workbook.Save("CustomDateAxisChart.xlsx");
    }
}
