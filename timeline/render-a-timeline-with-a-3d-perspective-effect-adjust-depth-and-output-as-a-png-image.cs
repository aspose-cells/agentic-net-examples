// Title: Generate a 3‑D column timeline chart with perspective in C# using Aspose.Cells and save it as an XLSX workbook
// Common Searches: C# Aspose.Cells create a 3D column chart for a project timeline | How to set perspective angle on a 3D chart with Aspose.Cells .NET | Adding date categories to a 3D column chart in Aspose.Cells | Saving an Aspose.Cells workbook that contains a 3D chart to an XLSX file | Configure chart title and data series for a timeline chart in Aspose.Cells C#
// Tags: Aspose.Cells 3D column chart creation | Aspose.Cells chart perspective setting | Aspose.Cells timeline data series definition | Aspose.Cells workbook export to XLSX | C# generate timeline chart with dates

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example demonstrates how to create a new Workbook, populate columns with dates and values, add a 3‑D column chart representing a timeline, assign the data series and category range, set a chart title, apply a 30‑degree perspective angle, and finally save the workbook containing the chart as 'timeline.xlsx' using Aspose.Cells for .NET.
class Timeline3D
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data: dates in column A and values in column B
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue(new DateTime(2023, 2, 1));
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["A4"].PutValue(new DateTime(2023, 3, 1));
            sheet.Cells["B4"].PutValue(180);
            sheet.Cells["A5"].PutValue(new DateTime(2023, 4, 1));
            sheet.Cells["B5"].PutValue(130);
            sheet.Cells["A6"].PutValue(new DateTime(2023, 5, 1));
            sheet.Cells["B6"].PutValue(170);

            // Add a 3‑D column chart to the worksheet
            // Parameters: chart type, upper‑left row, upper‑left column, lower‑right row, lower‑right column
            int chartIndex = sheet.Charts.Add(ChartType.Column3D, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data series (values) and category (dates)
            chart.NSeries.Add("B2:B6", true);
            chart.NSeries.CategoryData = "A2:A6";

            // Configure chart title
            chart.Title.Text = "Project Timeline (3‑D)";

            // Apply 3‑D perspective (available in Chart API)
            chart.Perspective = 30;   // Perspective angle (0‑90)

            // Note: Depth, RotationX, and RotationY properties are not available in the current Aspose.Cells version.
            // They can be omitted or replaced with other 3‑D settings if needed.

            // Save the workbook (which includes the chart) to an Excel file
            string outputPath = "timeline.xlsx";
            workbook.Save(outputPath);

            // Cleanup
            workbook.Dispose();
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
