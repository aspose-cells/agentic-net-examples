using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class ProgressBarChartExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data: task names in column A and progress percentages in column B
        sheet.Cells["A1"].PutValue("Task");
        sheet.Cells["B1"].PutValue("Progress");
        sheet.Cells["A2"].PutValue("Design");
        sheet.Cells["B2"].PutValue(0.25);   // 25%
        sheet.Cells["A3"].PutValue("Development");
        sheet.Cells["B3"].PutValue(0.55);   // 55%
        sheet.Cells["A4"].PutValue("Testing");
        sheet.Cells["B4"].PutValue(0.80);   // 80%
        sheet.Cells["A5"].PutValue("Deployment");
        sheet.Cells["B5"].PutValue(0.10);   // 10%

        // Add a 2‑D bar chart that will act as a progress bar
        int chartIndex = sheet.Charts.Add(ChartType.Bar, 7, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Add a series using the progress values (column B)
        // The Add method binds the series values; true indicates vertical (column) orientation
        chart.NSeries.Add("B2:B5", true);

        // Bind the category (task names) to the series
        chart.NSeries.CategoryData = "A2:A5";

        // Ensure the series is visible (not filtered out)
        chart.NSeries[0].IsFiltered = false;

        // Optional: format the values as percentages
        chart.NSeries[0].ValuesFormatCode = "0%";

        // Save the workbook to a file
        workbook.Save("ProgressBarChart.xlsx");
    }
}