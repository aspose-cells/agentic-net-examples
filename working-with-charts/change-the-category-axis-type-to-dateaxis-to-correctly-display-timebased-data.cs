// Title: Set a Date (TimeScale) Category Axis for a Line Chart using Aspose.Cells for .NET
// Description: Shows how to build a workbook, add a line chart with date‑value pairs, and change the CategoryAxis to TimeScale with month‑level base and major units before saving the file as XLSX.
// Keywords: Aspose.Cells | C# chart date axis | CategoryAxis TimeScale | line chart month intervals | BaseUnitScale months | TimeUnit.Months | .NET spreadsheet charting | date‑scaled axis
// Common Searches: Aspose.Cells set chart category axis to date | C# TimeScale axis Aspose.Cells example | how to use BaseUnitScale months in Aspose chart | line chart with monthly dates Aspose.Cells | configure CategoryAxis.CategoryType TimeScale .NET
// Developer Intent: Configure the chart’s X‑axis as a date‑scaled (TimeScale) axis so that time‑based data is displayed correctly.
// Use Cases: Monthly sales trend line chart with dates on the X‑axis. | Financial performance dashboard that requires precise time intervals. | Automated report generation where chart axes must reflect calendar months.
// AI Prompts: Generate C# code to set CategoryAxis.CategoryType to TimeScale and define month‑level BaseUnitScale in Aspose.Cells. | Explain how to adjust major unit settings for a date axis in an Aspose.Cells line chart. | Provide steps to format labels on a TimeScale axis after changing the category type in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDateAxisDemo
{
    // Shows how to build a workbook, add a line chart with date‑value pairs, and change the CategoryAxis to TimeScale with month‑level base and major units before saving the file as XLSX.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample date‑based data
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue(new DateTime(2024, 1, 1));
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue(new DateTime(2024, 2, 1));
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue(new DateTime(2024, 3, 1));
            sheet.Cells["B4"].PutValue(30);

            // Add a line chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and the category (X) axis
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Change the category axis type to TimeScale (date axis)
            chart.CategoryAxis.CategoryType = CategoryType.TimeScale;

            // Optional: define the base unit scale for better date spacing
            chart.CategoryAxis.BaseUnitScale = TimeUnit.Months;
            chart.CategoryAxis.MajorUnitScale = TimeUnit.Months;
            chart.CategoryAxis.MajorUnit = 1;

            // Save the workbook (lifecycle save rule)
            workbook.Save("DateAxisChart.xlsx", SaveFormat.Xlsx);
        }
    }
}
