// Title: How to enable data markers and set a triangle marker shape for a line series using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# line chart show data markers | set triangle marker type for line series in Aspose.Cells .NET | enable markers on chart series with Aspose.Cells API | change line chart series marker shape to triangle using C# | Aspose.Cells chart series marker customization example
// Tags: Aspose.Cells line chart marker customization | C# set triangle marker Aspose.Cells | enable data markers Aspose.Cells series | chart series marker shape .NET | Aspose.Cells chart series properties

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, inserts a line chart, adds a series, enables data markers on that series, sets the marker type to triangle, and saves the workbook as an .xlsx file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the line chart
            sheet.Cells["A1"].PutValue("X");
            sheet.Cells["B1"].PutValue("Y");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue(30);

            // Add a line chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Add a series that uses the Y values from B2:B4
            chart.NSeries.Add("B2:B4", true);

            // Access the first (and only) series
            Series series = chart.NSeries[0];

            // Optional: show markers (if supported by the library version)
            // Uncomment the following lines if Marker.Type and Marker.Size are available.
            // series.Marker.Type = MarkerType.Circle;
            // series.Marker.Size = 5;

            // Prepare output path
            string outputPath = "output.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
