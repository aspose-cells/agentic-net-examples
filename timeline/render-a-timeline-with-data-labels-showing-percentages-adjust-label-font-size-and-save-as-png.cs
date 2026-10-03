// Title: Generate a timeline line chart with percentage data labels, custom font size, and export it as PNG using Aspose.Cells for .NET
// AI Prompts: Create a line chart that uses dates on the X‑axis, display each point as a percentage label with a 12‑point font, and save the chart as a PNG image using Aspose.Cells in C#. | Add a timeline chart to a workbook, enable percentage data labels, adjust the label font size, and render the chart to a PNG file with Aspose.Cells for .NET.
// Common Searches: how to show percentage data labels on a line chart with Aspose.Cells C# | set custom font size for chart data labels in Aspose.Cells .NET | export Aspose.Cells chart as PNG image | create timeline chart from date column using Aspose.Cells
// Tags: line chart data labels percentage Aspose.Cells | custom data label font size C# Aspose.Cells | export chart to PNG Aspose.Cells .NET | timeline chart from dates Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example builds a workbook, fills it with date and value data, creates a line chart used as a timeline, enables percentage data labels with a 12‑point font, formats the labels as percentages, and saves the rendered chart as a PNG image.
class TimelineChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the timeline.
            // Column A: Dates (categories)
            // Column B: Values
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue(new DateTime(2023, 2, 1));
            sheet.Cells["B3"].PutValue(45);
            sheet.Cells["A4"].PutValue(new DateTime(2023, 3, 1));
            sheet.Cells["B4"].PutValue(25);
            sheet.Cells["A5"].PutValue(new DateTime(2023, 4, 1));
            sheet.Cells["B5"].PutValue(60);

            // Add a Line chart (used as a timeline representation).
            // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column.
            int upperLeftRow = 7;
            int upperLeftColumn = 0;
            int lowerRightRow = 25;
            int lowerRightColumn = 10;

            // Add returns the index of the newly created chart.
            int chartIndex = sheet.Charts.Add(ChartType.Line, upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart.
            // Category (X) values are in column A, series (Y) values are in column B.
            chart.NSeries.Add("B2:B5", false);
            chart.NSeries[0].XValues = "A2:A5";

            // Show percentages on data labels.
            chart.NSeries[0].DataLabels.ShowPercentage = true;

            // Adjust the font size of the data labels.
            chart.NSeries[0].DataLabels.Font.Size = 12; // Desired font size (e.g., 12 points).

            // Optionally, set the number format for percentages.
            chart.NSeries[0].DataLabels.NumberFormat = "0%";

            // Determine output file path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "timeline.png");

            // Ensure the directory exists.
            string? outputDir = Path.GetDirectoryName(outputPath);
            Directory.CreateDirectory(outputDir ?? Directory.GetCurrentDirectory());

            // Save the workbook as a PNG image (the chart will be rendered to the image).
            workbook.Save(outputPath, SaveFormat.Png);
            Console.WriteLine($"Timeline chart saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while creating the timeline chart:");
            Console.WriteLine(ex.Message);
        }
    }
}
