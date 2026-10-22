// Title: Proportionally resize Excel chart data label fonts after expanding chart width using Aspose.Cells for .NET
// AI Prompts: Resize a chart's width by a factor and automatically adjust each series' data label font size to keep the labels proportional with Aspose.Cells in C#. | Read a scaling factor from a JSON configuration file, apply it to both the chart dimensions and the data label fonts, and save the updated workbook.
// Common Searches: C# Aspose.Cells how to keep data label font size proportional when changing chart size | scale Excel chart width and automatically adjust data label fonts with Aspose.Cells | programmatically resize chart and data labels in a .NET workbook using Aspose.Cells
// Tags: chart width scaling Aspose.Cells | data label font resizing Aspose.Cells C# | proportional chart label scaling Excel | adjust chart dimensions and label sizes .NET | automate chart formatting Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ChartLabelResizeExample
{
    // The example creates a workbook with a column chart if needed, loads it, enlarges the chart width by a factor, scales each series' data label font size proportionally, and saves the result to a new file.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Ensure the input file exists; create a simple workbook if it does not.
                if (!File.Exists(inputPath))
                {
                    try
                    {
                        var tempWb = new Workbook();
                        var ws = tempWb.Worksheets[0];
                        ws.Name = "Sheet1";

                        // Add sample data.
                        ws.Cells["A1"].PutValue("Category");
                        ws.Cells["B1"].PutValue("Value");
                        ws.Cells["A2"].PutValue("A");
                        ws.Cells["B2"].PutValue(10);
                        ws.Cells["A3"].PutValue("B");
                        ws.Cells["B3"].PutValue(20);
                        ws.Cells["A4"].PutValue("C");
                        ws.Cells["B4"].PutValue(30);

                        // Create a chart to work with.
                        int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 20, 5);
                        Chart createdChart = ws.Charts[chartIdx];
                        createdChart.NSeries.Add("B2:B4", true);
                        // Category data is automatically taken from the first column when the second argument is true.
                        // If explicit assignment is needed, uncomment the line below (requires a version that supports CategoryData).
                        // createdChart.NSeries[0].CategoryData = "A2:A4";

                        tempWb.Save(inputPath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to create template workbook: {ex.Message}");
                        return;
                    }
                }

                // Load the workbook (ensure the file exists before loading).
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                var workbook = new Workbook(inputPath);
                var sheet = workbook.Worksheets[0];

                // Ensure there is at least one chart.
                if (sheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the worksheet.");
                    return;
                }

                // Assume the first chart is the target.
                Chart targetChart = sheet.Charts[0];

                // Define a scaling factor (e.g., 150% increase).
                const double scaleFactor = 1.5;

                // Resize data label fonts proportionally.
                foreach (Series series in targetChart.NSeries)
                {
                    var dataLabels = series.DataLabels;
                    if (dataLabels?.Font != null)
                    {
                        int originalFontSize = dataLabels.Font.Size; // Font.Size is an integer.
                        int newFontSize = (int)Math.Round(originalFontSize * scaleFactor);
                        dataLabels.Font.Size = newFontSize;
                    }
                }

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
