// Title: Set Category Axis to Date (TimeScale) Axis in Aspose.Cells for .NET Line Chart
// Description: Creates a workbook, fills column A with DateTime values and column B with numbers, adds a line chart, assigns the data ranges, then switches the chart's CategoryAxis to a TimeScale (date) axis and configures month‑based major ticks before saving the file.
// Keywords: Aspose.Cells date axis | CategoryAxis TimeScale .NET | line chart date axis Aspose | C# Aspose.Cells chart axis | TimeUnit.Months Aspose.Cells
// Common Searches: Aspose.Cells set X‑axis to date | TimeScale axis example C# | how to use CategoryType.TimeScale in Aspose.Cells | configure month intervals on chart axis Aspose | date based category axis Aspose.Cells .NET
// Developer Intent: Convert the chart’s category axis to a date (TimeScale) axis so the X‑axis reflects chronological spacing.
// Use Cases: Plot monthly sales data with accurate date spacing. | Display project milestones on a timeline chart. | Generate financial time‑series reports where the X‑axis uses a date scale.
// AI Prompts: Write C# code with Aspose.Cells that changes a chart’s category axis to a TimeScale axis and sets month intervals. | Explain how to adjust BaseUnitScale and MajorUnit for a date axis in Aspose.Cells for .NET. | Show how to format date labels on a TimeScale axis in an Aspose.Cells chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, fills column A with DateTime values and column B with numbers, adds a line chart, assigns the data ranges, then switches the chart's CategoryAxis to a TimeScale (date) axis and configures month‑based major ticks before saving the file.
class ChangeCategoryAxisToDateAxis
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate the worksheet with date‑based data
        sheet.Cells["A1"].PutValue("Date");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue(new DateTime(2024, 1, 1));
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue(new DateTime(2024, 2, 1));
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue(new DateTime(2024, 3, 1));
        sheet.Cells["B4"].PutValue(30);

        // Add a line chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the series and the category (X) axis
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Change the category axis type to TimeScale (date axis)
        chart.CategoryAxis.CategoryType = CategoryType.TimeScale;

        // Optional: define the base unit scale for better date spacing
        chart.CategoryAxis.BaseUnitScale = TimeUnit.Months;
        chart.CategoryAxis.MajorUnitScale = TimeUnit.Months;
        chart.CategoryAxis.MajorUnit = 1; // one month per major tick

        // Save the workbook
        workbook.Save("ChartWithDateAxis.xlsx");
    }
}
