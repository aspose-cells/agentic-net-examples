// Title: Aspose.Cells C# – Export Chart as PNG and Store Series & Axis Labels in JSON
// Description: Creates a workbook, adds sample data, builds a column chart, assigns axis titles, renders the chart to a PNG image, extracts each series name and the axis titles, then writes this metadata to a formatted JSON file placed alongside the image while also saving the original workbook.
// Keywords: Aspose.Cells | C# chart export | chart to PNG | chart metadata JSON | extract series names | axis titles | save chart image | Aspose.Cells example | Excel chart automation
// Common Searches: Aspose.Cells export chart as PNG with metadata | C# get chart series names and axis labels from Aspose.Cells | Save Aspose.Cells chart image and JSON description | How to serialize chart information to JSON in C# | Aspose.Cells chart to image and metadata file
// Developer Intent: Generate a PNG image of an Aspose.Cells chart and capture its series names and axis titles in a JSON file.
// Use Cases: Bundle chart images with machine‑readable descriptors for analytics pipelines. | Automate report creation where visual charts are exported and their labels are indexed for search. | Provide dynamic tooltips in web dashboards by reading series and axis information from JSON. | Archive Excel charts with accompanying metadata for compliance and documentation.
// AI Prompts: Write a C# method that reads the JSON metadata and returns a strongly‑typed object containing series, category axis title, value axis title, and chart title. | Modify the example to include the chart's main title in the JSON output and handle missing titles gracefully. | Add robust error handling so the PNG and JSON files are saved to a user‑specified folder, with a fallback to a temporary directory.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartMetadata
{
    // Creates a workbook, adds sample data, builds a column chart, assigns axis titles, renders the chart to a PNG image, extracts each series name and the axis titles, then writes this metadata to a formatted JSON file placed alongside the image while also saving the original workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (creation rule)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
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

                // Add a column chart (chart creation rule)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set data range for the chart
                chart.NSeries.Add("B2:C4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Optional: set titles for axes (metadata we will capture)
                chart.CategoryAxis.Title.Text = "Months";
                chart.ValueAxis.Title.Text = "Sales";

                // Save the chart as PNG image (use ToImage rule)
                string imagePath = "ChartImage.png";
                chart.ToImage(imagePath, ImageType.Png);

                // Gather chart metadata
                var metadata = new
                {
                    Series = GetSeriesInfo(chart),
                    CategoryAxisTitle = chart.CategoryAxis.Title.Text,
                    ValueAxisTitle = chart.ValueAxis.Title.Text,
                    ChartTitle = chart.Title.Text
                };

                // Serialize metadata to JSON
                string json = JsonSerializer.Serialize(metadata, new JsonSerializerOptions { WriteIndented = true });

                // Save JSON file alongside the PNG (use standard file I/O)
                string jsonPath = Path.ChangeExtension(imagePath, ".json");
                File.WriteAllText(jsonPath, json);

                // Save the workbook (save rule)
                string workbookPath = "ChartWorkbook.xlsx";
                workbook.Save(workbookPath);

                Console.WriteLine($"Chart image saved to: {Path.GetFullPath(imagePath)}");
                Console.WriteLine($"Metadata JSON saved to: {Path.GetFullPath(jsonPath)}");
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(workbookPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method to extract series names (ValuesData property not available in this version)
        private static List<object> GetSeriesInfo(Chart chart)
        {
            var list = new List<object>();
            for (int i = 0; i < chart.NSeries.Count; i++)
            {
                var series = chart.NSeries[i];
                // Series name may be empty; fallback to default naming if needed
                string name = !string.IsNullOrEmpty(series.Name) ? series.Name : $"Series{i + 1}";
                list.Add(new { Name = name });
            }
            return list;
        }
    }
}
