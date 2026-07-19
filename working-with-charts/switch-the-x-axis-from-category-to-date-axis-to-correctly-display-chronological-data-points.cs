// Title: C# – Set Line Chart X‑Axis to a Date (Time) Axis with Aspose.Cells
// Description: This example creates a workbook, adds a line chart, populates column A with DateTime values and column B with numeric data, then converts the X‑axis from a categorical axis to a time‑scale axis by setting CategoryAxis.CategoryType to CategoryType.TimeScale. It also demonstrates how to define base, major, and minor units (months and days) for precise scaling before saving the file as DateAxisChart.xlsx.
// Keywords: Aspose.Cells C# date axis | CategoryAxis TimeScale | line chart date axis .NET | Aspose.Cells chart time scale | configure chart axis units Aspose.Cells | CategoryType.TimeScale example | C# chart date axis | Aspose.Cells X axis date | time scale axis Aspose.Cells
// Common Searches: Aspose.Cells set X axis to date axis C# | how to use CategoryType.TimeScale in Aspose.Cells | C# line chart with time scale axis Aspose | configure major and minor units on date axis Aspose.Cells | convert categorical axis to time axis Aspose.Cells
// Developer Intent: Switch a chart’s X‑axis from categorical to date (time) scale and fine‑tune its tick intervals.
// Use Cases: Plot monthly sales data with dates spaced proportionally on the X‑axis. | Create a project timeline where each milestone date is accurately positioned. | Generate a financial trend chart with weekly minor ticks for detailed analysis.
// AI Prompts: Show me C# code to change a line chart X‑axis to a date axis using Aspose.Cells and set month‑based major ticks. | Provide an Aspose.Cells example that creates a line chart with DateTime categories and custom major/minor units. | How do I configure CategoryAxis.CategoryType to TimeScale and define base, major, and minor units in Aspose.Cells for .NET?

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDateAxisDemo
{
    // This example creates a workbook, adds a line chart, populates column A with DateTime values and column B with numeric data, then converts the X‑axis from a categorical axis to a time‑scale axis by setting CategoryAxis.CategoryType to CategoryType.TimeScale. It also demonstrates how to define base, major, and minor units (months and days) for precise scaling before saving the file as DateAxisChart.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data: dates in column A, numeric values in column B
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue(new DateTime(2024, 1, 1));
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue(new DateTime(2024, 2, 1));
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue(new DateTime(2024, 3, 1));
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["A5"].PutValue(new DateTime(2024, 4, 1));
            sheet.Cells["B5"].PutValue(25);
            sheet.Cells["A6"].PutValue(new DateTime(2024, 5, 1));
            sheet.Cells["B6"].PutValue(35);

            // Add a line chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Line, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series (values) and the category (dates)
            chart.NSeries.Add("B2:B6", true);
            chart.NSeries.CategoryData = "A2:A6";

            // Switch the X axis from a categorical axis to a date (time) axis
            chart.CategoryAxis.CategoryType = CategoryType.TimeScale;

            // Optional: define the base unit and major/minor units for better scaling
            chart.CategoryAxis.BaseUnitScale = TimeUnit.Months;   // Base unit = months
            chart.CategoryAxis.MajorUnitScale = TimeUnit.Months; // Major ticks every month
            chart.CategoryAxis.MajorUnit = 1;
            chart.CategoryAxis.MinorUnitScale = TimeUnit.Days;   // Minor ticks every day
            chart.CategoryAxis.MinorUnit = 7;                    // One week

            // Add titles for clarity
            chart.Title.Text = "Sample Chart with Date Axis";
            chart.CategoryAxis.Title.Text = "Date";
            chart.ValueAxis.Title.Text = "Value";

            // Save the workbook
            workbook.Save("DateAxisChart.xlsx");
        }
    }
}
