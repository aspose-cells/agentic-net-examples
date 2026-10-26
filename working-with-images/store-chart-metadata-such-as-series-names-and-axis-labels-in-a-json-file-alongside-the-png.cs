// Title: Export an Aspose.Cells chart to PNG and write its series names and axis labels into a JSON file using C#
// Common Searches: how to export an Aspose.Cells chart as a PNG image in C# | C# code to retrieve series names from an Aspose.Cells chart | save chart axis titles to a JSON file with Aspose.Cells | write chart metadata JSON alongside chart image using Aspose.Cells | Aspose.Cells example for rendering chart and serializing metadata
// Tags: Aspose.Cells chart PNG export | Aspose.Cells serialize chart metadata to JSON | C# extract chart series names Aspose.Cells | Aspose.Cells chart axis titles JSON | save chart image and metadata together .NET

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The sample creates a workbook, populates data, adds a column chart, sets titles, renders the chart as a PNG file, collects series names and axis labels into a ChartMetadata object, serializes this object to formatted JSON, and saves the JSON file alongside the PNG image using Aspose.Cells in C#.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet's cells
            var workbook = new Workbook();
            var cells = workbook.Worksheets[0].Cells;

            // Populate sample data for the chart
            cells["A1"].PutValue("Month");
            cells["B1"].PutValue("Series1");
            cells["C1"].PutValue("Series2");
            cells["A2"].PutValue("Jan");
            cells["A3"].PutValue("Feb");
            cells["A4"].PutValue("Mar");
            cells["B2"].PutValue(10);
            cells["B3"].PutValue(20);
            cells["B4"].PutValue(30);
            cells["C2"].PutValue(15);
            cells["C3"].PutValue(25);
            cells["C4"].PutValue(35);

            // Add a column chart to the worksheet
            var sheet = workbook.Worksheets[0];
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            var chart = sheet.Charts[chartIndex];

            // Set chart title and axis titles
            chart.Title.Text = "Sales Report";
            chart.CategoryAxis.Title.Text = "Month";
            chart.ValueAxis.Title.Text = "Revenue";

            // Add series and assign names
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries[0].Name = "Series1";
            chart.NSeries.Add("C2:C4", true);
            chart.NSeries[1].Name = "Series2";

            // Save the chart as a PNG file
            var imgOptions = new ImageOrPrintOptions
            {
                SaveFormat = SaveFormat.Png
            };
            string pngPath = "Chart.png";
            chart.ToImage(pngPath, imgOptions);

            // Gather chart metadata
            var metadata = new ChartMetadata
            {
                SeriesNames = new List<string>(),
                XAxisLabel = chart.CategoryAxis.Title.Text,
                YAxisLabel = chart.ValueAxis.Title.Text
            };

            // Collect series names
            for (int i = 0; i < chart.NSeries.Count; i++)
            {
                var series = chart.NSeries[i];
                metadata.SeriesNames.Add(series.Name);
            }

            // Serialize metadata to JSON and save alongside the PNG
            string jsonPath = "ChartMetadata.json";
            string json = JsonSerializer.Serialize(metadata, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(jsonPath, json);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}

// Helper class to hold chart metadata
public class ChartMetadata
{
    public List<string> SeriesNames { get; set; } = new List<string>();
    public string XAxisLabel { get; set; } = string.Empty;
    public string YAxisLabel { get; set; } = string.Empty;
}
