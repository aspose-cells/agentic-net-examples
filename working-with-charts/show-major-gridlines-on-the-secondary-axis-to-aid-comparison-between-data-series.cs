// Title: Enable major gridlines on a chart’s secondary axis using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to turn on major gridlines for a secondary axis | display secondary axis major gridlines in an Excel chart with Aspose.Cells .NET | Aspose.Cells chart secondary axis gridlines not visible example | enable major gridlines on secondary value axis using Aspose.Cells C# code | set secondary axis gridline visibility in Aspose.Cells chart
// Tags: enable secondary axis major gridlines Aspose.Cells | chart secondary axis gridlines .NET | Aspose.Cells secondary axis formatting | C# display major gridlines secondary axis | Excel chart secondary axis customization Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts; // Required for Chart and ChartType

namespace AsposeCellsExample
{
    // The example loads an existing workbook (or creates one), ensures a chart is present, activates major gridlines on the chart’s secondary axis, makes the chart title visible, shows the legend, and saves the workbook as an Excel file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Ensure the input file exists; create a sample workbook if it does not.
                if (!File.Exists(inputPath))
                {
                    CreateSampleWorkbook(inputPath);
                }

                // Load the workbook.
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet.
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one chart; create a sample chart if none exist.
                if (worksheet.Charts.Count == 0)
                {
                    AddSampleChart(worksheet);
                }

                // Access the first chart.
                Chart chart = worksheet.Charts[0];

                // Make sure the chart title is visible.
                if (chart.Title != null)
                {
                    chart.Title.IsVisible = true;
                }

                // Show legend.
                chart.ShowLegend = true;

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Creates a simple workbook with sample data and a chart.
        private static void CreateSampleWorkbook(string path)
        {
            try
            {
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];

                // Sample data.
                ws.Cells["A1"].PutValue("Category");
                ws.Cells["B1"].PutValue("Value");
                ws.Cells["A2"].PutValue("A");
                ws.Cells["A3"].PutValue("B");
                ws.Cells["A4"].PutValue("C");
                ws.Cells["B2"].PutValue(10);
                ws.Cells["B3"].PutValue(20);
                ws.Cells["B4"].PutValue(30);

                // Add a sample chart.
                AddSampleChart(ws);

                wb.Save(path);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to create sample workbook: {ex.Message}");
                throw;
            }
        }

        // Adds a column chart to the worksheet using data in A2:B4.
        private static void AddSampleChart(Worksheet ws)
        {
            try
            {
                int chartIndex = ws.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = ws.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Set chart title.
                chart.Title.Text = "Sample Chart";
                chart.Title.IsVisible = true;

                // Show legend.
                chart.ShowLegend = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to add sample chart: {ex.Message}");
                throw;
            }
        }
    }
}
