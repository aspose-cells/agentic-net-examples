// Title: Convert the X axis from category to date (time scale) axis in an Aspose.Cells line chart using C#
// Common Searches: Aspose.Cells C# set line chart X axis to time scale | change category axis to date axis in Aspose.Cells chart programmatically | C# Aspose.Cells line chart with daily date axis | how to use CategoryType.TimeScale for chart axis in Aspose.Cells .NET | set BaseUnitScale to Days for date axis in Aspose.Cells line chart | Aspose.Cells chart axis formatting for chronological data
// Tags: Aspose.Cells line chart date axis | CategoryAxis CategoryType.TimeScale usage | BaseUnitScale Days setting Aspose.Cells | C# chronological data chart Aspose.Cells | time scale X axis Aspose.Cells .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, fills column A with sequential dates and column B with sales figures, adds a line chart, assigns the data ranges, switches the chart's CategoryAxis to a time‑scale (date) axis using CategoryType.TimeScale, sets the base unit to days, titles the axes and chart, recalculates, and saves the file as an XLSX workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with chronological data (dates in column A, values in column B)
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Sales");

            DateTime startDate = new DateTime(2023, 1, 1);
            for (int i = 0; i < 10; i++)
            {
                // Column A: Date values
                sheet.Cells[i + 1, 0].PutValue(startDate.AddDays(i));
                // Column B: Some numeric data
                sheet.Cells[i + 1, 1].PutValue(100 + i * 15);
            }

            // Add a line chart
            int chartIndex = sheet.Charts.Add(ChartType.Line, 12, 0, 30, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series (values) and the category data (dates)
            chart.NSeries.Add("B2:B11", true);          // Values
            chart.NSeries.CategoryData = "A2:A11";      // Dates

            // Switch the X (category) axis to a date (time) axis
            Axis categoryAxis = chart.CategoryAxis;
            categoryAxis.CategoryType = CategoryType.TimeScale;   // Correct enum value for date axis
            categoryAxis.BaseUnitScale = TimeUnit.Days;           // Optional: define base unit scale

            // Adjust visual properties
            categoryAxis.Title.Text = "Date";
            chart.ValueAxis.Title.Text = "Sales";
            chart.Title.Text = "Sales Over Time";

            // Recalculate the chart to apply changes
            chart.Calculate();

            // Save the workbook
            string outFile = "DateAxisChart.xlsx";
            workbook.Save(outFile, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
