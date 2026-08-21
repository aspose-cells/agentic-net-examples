// Title: Set X‑Axis to Date (TimeScale) in a Line Chart with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills column A with DateTime values and column B with numbers, adds a line chart, links the series, and configures the CategoryAxis to TimeScale with month‑based base and major units before saving the file.
// Keywords: Aspose.Cells | C# chart date axis | TimeScale CategoryAxis | line chart with dates | BaseUnitScale TimeUnit.Months | set major unit Aspose.Cells | date‑scaled chart .NET | configure X axis as date | chart axis time scale example
// Common Searches: Aspose.Cells set X axis to date | CategoryAxis TimeScale C# example | How to use TimeUnit.Months in Aspose.Cells chart | Create line chart with date axis Aspose.Cells | Configure chart axis as time scale .NET
// Developer Intent: The developer wants the chart’s X‑axis to interpret category values as dates so that time‑based data is plotted correctly, with control over the axis interval units.
// Use Cases: Plot monthly sales figures with a date‑based X‑axis for clear trend visualization. | Generate a project timeline chart where milestones are positioned by calendar dates. | Export financial reports that show quarterly performance using a time‑scaled axis for better spacing.
// AI Prompts: Show C# code to set CategoryAxis.CategoryType to TimeScale and define month intervals in Aspose.Cells. | Explain how to format X‑axis labels as short dates after enabling a TimeScale axis. | Provide steps to add multiple series to a date‑scaled line chart using Aspose.Cells. | Give an example of adjusting BaseUnitScale and MajorUnit for a chart with daily data points.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, fills column A with DateTime values and column B with numbers, adds a line chart, links the series, and configures the CategoryAxis to TimeScale with month‑based base and major units before saving the file.
class SetXAxisDateAxis
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate worksheet with date (X) and numeric (Y) data
        worksheet.Cells["A1"].PutValue("Date");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue(new DateTime(2024, 1, 1));
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["A3"].PutValue(new DateTime(2024, 2, 1));
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["A4"].PutValue(new DateTime(2024, 3, 1));
        worksheet.Cells["B4"].PutValue(30);

        // Add a line chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Line, 5, 0, 25, 10);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data source for the chart
        chart.NSeries.Add("B2:B4", true);          // Y‑values
        chart.NSeries.CategoryData = "A2:A4";      // X‑values (dates)

        // Configure the X‑axis (category axis) to treat values as dates
        chart.CategoryAxis.CategoryType = CategoryType.TimeScale;

        // Optional: define the base unit and major unit for clearer spacing
        chart.CategoryAxis.BaseUnitScale = TimeUnit.Months;
        chart.CategoryAxis.MajorUnitScale = TimeUnit.Months;
        chart.CategoryAxis.MajorUnit = 1;

        // Save the workbook
        workbook.Save("ChartWithDateAxis.xlsx");
    }
}
