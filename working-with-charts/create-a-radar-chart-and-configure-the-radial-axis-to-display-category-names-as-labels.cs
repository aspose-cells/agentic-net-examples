// Title: How to create a radar chart with custom radial axis labels from worksheet data using Aspose.Cells for .NET (C#)
// Common Searches: aspnetc# aspose.cells create radar chart with category axis labels from column A | set radial axis major unit to 1 for Aspose.Cells radar chart in C# | bind radar chart series to cell ranges B2:B6 and C2:C6 using Aspose.Cells .NET | add radar chart to worksheet and configure category axis labels in Aspose.Cells | save Excel workbook containing radar chart with custom radial labels using Aspose.Cells
// Tags: Aspose.Cells radar chart category axis configuration | C# set radial axis major unit Aspose.Cells | Aspose.Cells bind chart series to cell ranges | Create Excel radar chart with Aspose.Cells .NET | Aspose.Cells save workbook with radar chart

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace RadarChartExample
{
    // The example creates a new Workbook, fills column A with category names and columns B‑C with two data series, adds a Radar chart to the first worksheet, links the series to ranges B2:B6 and C2:C6, assigns the radial (category) axis labels from A2:A6, disables automatic scaling and sets the major unit to 1 so each label appears, adds a chart title, ensures the output directory exists, and saves the file as RadarChartWithRadialLabels.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Access the first worksheet
                var sheet = workbook.Worksheets[0];

                // Populate sample data for the radar chart
                // Column A: Category names (radial axis labels)
                // Columns B and C: Series values
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Series 1");
                sheet.Cells["C1"].PutValue("Series 2");

                sheet.Cells["A2"].PutValue("Alpha");
                sheet.Cells["A3"].PutValue("Beta");
                sheet.Cells["A4"].PutValue("Gamma");
                sheet.Cells["A5"].PutValue("Delta");
                sheet.Cells["A6"].PutValue("Epsilon");

                sheet.Cells["B2"].PutValue(4);
                sheet.Cells["B3"].PutValue(7);
                sheet.Cells["B4"].PutValue(3);
                sheet.Cells["B5"].PutValue(6);
                sheet.Cells["B6"].PutValue(5);

                sheet.Cells["C2"].PutValue(5);
                sheet.Cells["C3"].PutValue(6);
                sheet.Cells["C4"].PutValue(4);
                sheet.Cells["C5"].PutValue(7);
                sheet.Cells["C6"].PutValue(3);

                // Add a radar chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Radar, 8, 0, 22, 10);
                var chart = sheet.Charts[chartIndex];

                // Set the data source for the chart (including category names)
                chart.NSeries.Add("B2:B6", true);
                chart.NSeries.Add("C2:C6", true);

                // Assign the category (radial) axis labels from column A
                chart.NSeries.CategoryData = "A2:A6";

                // Configure the radial (category) axis to display the category names
                var radialAxis = chart.CategoryAxis;
                radialAxis.IsAutomaticMajorUnit = false;          // Disable automatic scaling
                radialAxis.MajorUnit = 1;                         // Ensure each category gets a label
                // Note: TickLabelPosition is optional; omitted for compatibility

                // Optional: set a title for clarity
                chart.Title.Text = "Radar Chart with Category Labels";

                // Define output file path
                string outputPath = "RadarChartWithRadialLabels.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the radar chart:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
