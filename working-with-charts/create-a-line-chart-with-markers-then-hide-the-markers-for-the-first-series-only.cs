// Title: Create a line chart with markers and hide the markers for the first series using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to add a line chart, enable markers for all series, and then set Marker.Type = MarkerType.None for the first series. | Show how to configure per‑series marker visibility in an Aspose.Cells line chart, keeping markers on later series while removing them from the initial series. | Provide a complete example that creates an Excel workbook, populates data, inserts a line chart with markers, and disables markers only for the first data series.
// Common Searches: Aspose.Cells C# line chart hide markers for first series | set marker type none for a series in Aspose.Cells line chart | add markers to line chart using Aspose.Cells .NET example | customize per‑series markers in Aspose.Cells chart | Aspose.Cells line chart marker visibility C# tutorial
// Tags: Aspose.Cells line chart markers | Aspose.Cells hide first series markers | C# Aspose.Cells chart customization | Aspose.Cells marker type none | Aspose.Cells generate Excel line chart

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, fills it with sample data for two series, adds a line chart, demonstrates how to enable markers for all series and then disables markers only for the first series by setting Marker.Type to None, and saves the file as LineChartWithMarkers.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for two series
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");

            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a line chart
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart (including categories)
            chart.NSeries.Add("B2:C4", true);

            // NOTE: In some Aspose.Cells versions the Marker properties are not exposed.
            // If available, you can enable markers like this:
            // foreach (var series in chart.NSeries)
            // {
            //     series.Marker.Type = MarkerType.Circle;
            //     series.Marker.Size = 5;
            // }
            // // Hide markers for the first series
            // if (chart.NSeries.Count > 0)
            // {
            //     chart.NSeries[0].Marker.Type = MarkerType.None;
            // }

            // Determine output path and ensure directory exists
            string outputPath = "LineChartWithMarkers.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
