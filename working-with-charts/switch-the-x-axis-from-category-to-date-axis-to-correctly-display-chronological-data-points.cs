// Title: Switch Chart X‑Axis to a Date (Time‑Scale) Axis using Aspose.Cells for .NET (C#)
// Description: Creates a workbook with date and numeric data, adds a line chart, and converts the X‑axis from a categorical axis to a time‑scale (date) axis by setting CategoryAxis.CategoryType to TimeScale, configuring base, major and minor units, and formatting labels as "MMM yyyy" before saving the file.
// Keywords: Aspose.Cells chart date axis | C# time scale axis | CategoryAxis TimeScale example | Aspose.Cells line chart X axis | format X axis labels Aspose.Cells | .NET chart date axis configuration | switch X axis to date Aspose.Cells
// Common Searches: Aspose.Cells change X axis to date axis | C# set chart category axis to time scale | Aspose.Cells line chart monthly major unit | how to format chart X axis labels MMM yyyy Aspose.Cells | Aspose.Cells date axis minor unit days
// Developer Intent: Transform a chart’s X‑axis from a categorical axis to a date (time‑scale) axis so data points are plotted chronologically.
// Use Cases: Plot monthly sales data with dates on a time‑scale X‑axis for financial dashboards. | Display project milestones on a weekly‑minor‑unit, monthly‑major‑unit axis in status reports. | Generate Excel charts for inventory trends where the X‑axis shows month and year labels.
// AI Prompts: Show C# code with Aspose.Cells to convert a chart’s X‑axis to a time‑scale axis and set monthly major units and weekly minor units. | Provide an example of formatting X‑axis labels as "MMM yyyy" in an Aspose.Cells line chart. | Explain how to switch a chart’s category axis from CategoryType.Category to CategoryType.TimeScale and adjust base, major, and minor unit scales in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook with date and numeric data, adds a line chart, and converts the X‑axis from a categorical axis to a time‑scale (date) axis by setting CategoryAxis.CategoryType to TimeScale, configuring base, major and minor units, and formatting labels as "MMM yyyy" before saving the file.
    public class SwitchXAxisToDateAxis
    {
        // Entry point required for console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully: SwitchXAxisToDateAxis.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample date and numeric data
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

            // Add a line chart
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and the category (X) axis
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Switch the X axis from a categorical axis to a date (time) axis
            chart.CategoryAxis.CategoryType = CategoryType.TimeScale;   // Use time scale
            chart.CategoryAxis.BaseUnitScale = TimeUnit.Months;        // Base unit (optional)
            chart.CategoryAxis.MajorUnitScale = TimeUnit.Months;       // Major grid every month
            chart.CategoryAxis.MajorUnit = 1;                         // One month per major unit
            chart.CategoryAxis.MinorUnitScale = TimeUnit.Days;        // Minor grid daily
            chart.CategoryAxis.MinorUnit = 7;                         // One week per minor unit

            // Optional: format the X‑axis labels to show month/year
            chart.NSeries[0].XValuesFormatCode = "mmm yyyy";

            // Save the workbook
            workbook.Save("SwitchXAxisToDateAxis.xlsx");
        }
    }
}
