// Title: How to hide the horizontal axis major gridlines in a scatter chart with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide X-axis major gridlines in a scatter plot | remove horizontal axis gridlines from Excel scatter chart using Aspose.Cells | disable category axis major gridlines in .NET Excel chart programmatically | C# code to turn off scatter chart gridlines with Aspose.Cells | cleaner scatter chart appearance Aspose.Cells hide axis gridlines
// Tags: Aspose.Cells scatter chart major gridlines | C# hide category axis gridlines | Excel chart formatting with Aspose.Cells | disable horizontal axis gridlines .NET | chart appearance customization Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds X/Y data, inserts a scatter chart, assigns the data series, and hides the horizontal (category) axis major gridlines by setting Chart.CategoryAxis.MajorGridLines.IsVisible to false, then saves the file.
class HideHorizontalAxisGridlines
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the scatter plot (X in column A, Y in column B).
            sheet.Cells["A1"].PutValue("X");
            sheet.Cells["B1"].PutValue("Y");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue(2);
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue(4);
            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue(6);
            sheet.Cells["A5"].PutValue(4);
            sheet.Cells["B5"].PutValue(8);

            // Add a scatter chart to the worksheet.
            int chartIndex = sheet.Charts.Add(ChartType.Scatter, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart.
            // Y values (Values) from B2:B5, X values (Category) from A2:A5.
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries[0].XValues = "A2:A5";

            // Hide the horizontal axis (X axis) major gridlines for a cleaner appearance.
            chart.CategoryAxis.MajorGridLines.IsVisible = false;

            // Optionally hide minor gridlines.
            // chart.CategoryAxis.MinorGridLines.IsVisible = false;

            // Define output file path.
            string outputPath = "ScatterPlot_NoHorizontalGridlines.xlsx";

            // Ensure the output directory exists.
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
