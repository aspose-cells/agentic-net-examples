// Title: Save an Aspose.Cells column chart as PNG and generate a matching JSON file with chart title, axis labels, and series references in C#
// AI Prompts: Generate C# code that creates a workbook, adds a column chart, exports the chart to a PNG file, and writes the chart's title, category axis title, value axis title, and each series name with its data range to a JSON file beside the image. | Write a method that receives an Aspose.Cells Chart object, extracts its metadata (title, axis titles, series names and value references), and serializes the information to an indented JSON file using System.Text.Json. | Implement robust error handling and automatic output folder creation when saving a chart image and its metadata JSON with Aspose.Cells in a C# console application.
// Common Searches: how to export Aspose.Cells chart to PNG and also save chart metadata as JSON in C# | C# Aspose.Cells extract series name and data range from chart | save chart image and JSON file with same base name using Aspose.Cells | Aspose.Cells create output directory automatically when exporting chart | serialize Aspose.Cells chart titles and axis labels to JSON file
// Tags: Aspose.Cells chart PNG export with metadata JSON | extract chart series data range Aspose.Cells | serialize chart titles axis labels to JSON C# | auto-create output folder Aspose.Cells export | column chart generation Aspose.Cells C#

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// This C# example creates a workbook, adds sample data, builds a column chart, exports the chart as a PNG image, extracts the chart title, axis titles, and each series name with its cell range, then writes this information to an indented JSON file that shares the PNG's filename, while ensuring the output directory exists.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["C1"].PutValue("Series 2");

            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title and axis titles
            chart.Title.Text = "Sales Overview";
            chart.CategoryAxis.Title.Text = "Month";
            chart.ValueAxis.Title.Text = "Revenue";

            // Add series to the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries[0].Name = "Series 1";

            chart.NSeries.Add("C2:C4", true);
            chart.NSeries[1].Name = "Series 2";

            // Set the category (X) axis labels
            chart.NSeries.CategoryData = "A2:A4";

            // Export the chart as a PNG image
            string imagePath = "Chart.png";

            try
            {
                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(imagePath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Configure image export options
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    // Use ImageType instead of the obsolete SaveFormat
                    ImageType = ImageType.Png,
                    OnePagePerSheet = true
                };

                // Save the chart image to file (parameter order: filePath, options)
                chart.ToImage(imagePath, imgOptions);
            }
            catch (Exception imgEx)
            {
                Console.WriteLine($"Image export error: {imgEx.Message}");
                throw;
            }

            // Gather chart metadata
            var seriesMetadata = new List<object>();
            foreach (var series in chart.NSeries)
            {
                seriesMetadata.Add(new
                {
                    Name = series.Name,
                    // Values are stored as a string reference (e.g., "B2:B4")
                    ValuesReference = series.Values
                });
            }

            var chartMetadata = new
            {
                Title = chart.Title.Text,
                CategoryAxisTitle = chart.CategoryAxis.Title.Text,
                ValueAxisTitle = chart.ValueAxis.Title.Text,
                Series = seriesMetadata
            };

            // Serialize metadata to JSON
            string json = JsonSerializer.Serialize(chartMetadata, new JsonSerializerOptions { WriteIndented = true });

            // Save JSON file alongside the PNG
            string jsonPath = Path.ChangeExtension(imagePath, ".json");
            File.WriteAllText(jsonPath, json);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
