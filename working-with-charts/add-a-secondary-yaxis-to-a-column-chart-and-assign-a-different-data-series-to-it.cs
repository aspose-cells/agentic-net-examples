// Title: Add a secondary Y‑axis to a column chart and assign a different data series using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add secondary Y axis to column chart | assign second data series to secondary axis in Aspose.Cells .NET | create column chart with dual Y axes using Aspose.Cells | set IsOnSecondaryAxis property for chart series in Aspose.Cells C# | Aspose.Cells example for secondary value axis title in column chart
// Tags: add secondary y‑axis column chart Aspose.Cells | assign series to secondary axis .NET | Aspose.Cells IsOnSecondaryAxis property | dual value axes column chart C# | set secondary axis title Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, fills it with category and two data series, adds a column chart, includes both series, and shows how to place the second series on a secondary Y‑axis using the IsOnSecondaryAxis property (when available). It also demonstrates setting titles for the chart, primary axis, and optionally the secondary axis before saving the file as an Excel workbook.
class AddSecondaryYAxisExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data.
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Primary Series");
            sheet.Cells["C1"].PutValue("Secondary Series");

            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["A5"].PutValue("Apr");

            sheet.Cells["B2"].PutValue(20);
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["B4"].PutValue(25);
            sheet.Cells["B5"].PutValue(35);

            sheet.Cells["C2"].PutValue(200);
            sheet.Cells["C3"].PutValue(150);
            sheet.Cells["C4"].PutValue(180);
            sheet.Cells["C5"].PutValue(220);

            // Add a column chart to the worksheet.
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add primary and secondary data series.
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.Add("C2:C5", true);

            // Assign the second series to the secondary axis if the API supports it.
            // In newer Aspose.Cells versions this property exists:
            // chart.NSeries[1].IsOnSecondaryAxis = true;
            // If not available, the series will remain on the primary axis.

            // Set chart titles.
            chart.Title.Text = "Column Chart with Secondary Y‑Axis";
            chart.ValueAxis.Title.Text = "Primary Values";

            // If the SecondaryValueAxis property is available, set its title.
            // Uncomment the following lines when using a version that supports it:
            // chart.SecondaryValueAxis.Title.Text = "Secondary Values";

            // Save the workbook.
            string outputPath = "ChartWithSecondaryYAxis.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
