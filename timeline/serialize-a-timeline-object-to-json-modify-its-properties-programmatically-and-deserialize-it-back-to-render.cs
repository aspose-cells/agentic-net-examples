// Title: Serialize and modify Aspose.Cells timeline chart data to JSON in C# and reload into workbook
// AI Prompts: Generate C# code that creates a line chart with Aspose.Cells, extracts its title and data series into a TimelineData object, serializes it to indented JSON, then reads modified JSON and updates the chart title. | Write a C# snippet that uses System.Text.Json to change the title of an Aspose.Cells timeline chart by editing the JSON representation and applying the deserialized values back to the worksheet.
// Common Searches: C# Aspose.Cells how to export chart series to JSON | update Aspose.Cells chart title from JSON file | serialize Aspose.Cells timeline chart data using System.Text.Json | deserialize modified chart JSON and apply to workbook in C# | Aspose.Cells example for JSON round‑trip of chart properties
// Tags: Aspose.Cells chart JSON serialization | C# Aspose.Cells timeline chart update | System.Text.Json chart property deserialization | Aspose.Cells export chart data to JSON | modify Aspose.Cells chart title programmatically

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

// Demonstrates creating a line chart in Aspose.Cells, capturing its title and data points into a TimelineData object, serializing the object to indented JSON, programmatically altering the JSON (e.g., changing the title), deserializing the modified JSON, applying the updated properties back to the chart, and saving the workbook.
class TimelineData
{
    public string Title { get; set; } = string.Empty;
    public List<string> SeriesNames { get; set; } = new List<string>();
    public List<DateTime> XValues { get; set; } = new List<DateTime>();
    public List<double> YValues { get; set; } = new List<double>();
}

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate sample data for the timeline (date column + numeric column)
            ws.Cells["A1"].PutValue("Date");
            ws.Cells["B1"].PutValue("Value");
            DateTime start = new DateTime(2023, 1, 1);
            for (int i = 0; i < 10; i++)
            {
                ws.Cells[1 + i, 0].PutValue(start.AddDays(i));
                ws.Cells[1 + i, 1].PutValue(i * 10);
            }

            // Add a line chart as a substitute for a timeline chart
            int chartIdx = ws.Charts.Add(ChartType.Line, 5, 0, 20, 10);
            Chart timeline = ws.Charts[chartIdx];

            // Set the data source for the chart
            timeline.NSeries.Add("B2:B11", true);          // Values
            timeline.NSeries.CategoryData = "A2:A11";     // Dates
            timeline.Title.Text = "Sales Timeline";

            // Serialize chart properties to JSON
            TimelineData data = new TimelineData
            {
                Title = timeline.Title.Text,
                SeriesNames = new List<string> { "Value" }
            };
            for (int i = 0; i < 10; i++)
            {
                data.XValues.Add(ws.Cells[1 + i, 0].DateTimeValue);
                data.YValues.Add(ws.Cells[1 + i, 1].DoubleValue);
            }

            var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
            string json = JsonSerializer.Serialize(data, jsonOptions);
            Console.WriteLine("Serialized JSON:");
            Console.WriteLine(json);

            // Simulate external modification
            data.Title = "Updated Sales Timeline";
            string modifiedJson = JsonSerializer.Serialize(data, jsonOptions);
            Console.WriteLine("Modified JSON:");
            Console.WriteLine(modifiedJson);

            // Deserialize JSON back to object
            TimelineData? deserialized = JsonSerializer.Deserialize<TimelineData>(modifiedJson, jsonOptions);

            // Apply deserialized properties to the chart
            if (deserialized != null)
            {
                timeline.Title.Text = deserialized.Title;
                // Additional updates to series data could be added here if needed
            }

            // Save the workbook
            string outputPath = "TimelineOutput.xlsx";
            try
            {
                string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? "";
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
