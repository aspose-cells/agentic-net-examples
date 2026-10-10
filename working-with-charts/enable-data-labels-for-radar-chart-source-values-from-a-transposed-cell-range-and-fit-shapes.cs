// Title: Create a radar chart from a transposed cell range, show value and category data labels, and auto‑fit the chart size using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to add a radar chart whose series are built from a transposed range (rows become series) and turn on both value and category data labels for every series. | Provide C# that adjusts the radar chart’s dimensions to fit its contents automatically after the series are added, using the appropriate Aspose.Cells chart‑auto‑fit method.
// Common Searches: Aspose.Cells C# enable value and category data labels on radar chart series | how to use a transposed range as source for radar chart in Aspose.Cells | C# auto‑adjust radar chart size after adding series with Aspose.Cells | create radar chart from rows as series Aspose.Cells .NET example | display data labels for each series in Aspose.Cells radar chart
// Tags: radar chart transposed range Aspose.Cells | data labels series Aspose.Cells C# | auto‑fit chart dimensions Aspose.Cells | Aspose.Cells radar chart series configuration | C# generate radar chart from worksheet data

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace RadarChartExample
{
    // The program creates a new workbook, fills it with sample data, adds a radar chart using the transposed range B2:D5, enables both value and category data labels for each series, and saves the workbook as RadarChart_Output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (4 categories, 3 series)
                sheet.Cells["A2"].PutValue("Category 1");
                sheet.Cells["A3"].PutValue("Category 2");
                sheet.Cells["A4"].PutValue("Category 3");
                sheet.Cells["A5"].PutValue("Category 4");

                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["C2"].PutValue(20);
                sheet.Cells["D2"].PutValue(30);
                sheet.Cells["B3"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["D3"].PutValue(35);
                sheet.Cells["B4"].PutValue(20);
                sheet.Cells["C4"].PutValue(30);
                sheet.Cells["D4"].PutValue(40);
                sheet.Cells["B5"].PutValue(25);
                sheet.Cells["C5"].PutValue(35);
                sheet.Cells["D5"].PutValue(45);

                // Add a radar chart
                int chartIndex = sheet.Charts.Add(ChartType.Radar, 7, 0, 27, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the chart title
                chart.Title.Text = "Radar Chart with Transposed Data";

                // Add series using a transposed range (rows as series)
                chart.NSeries.Add("B2:D5", false);

                // Enable data labels for all series
                foreach (Series series in chart.NSeries)
                {
                    series.DataLabels.ShowValue = true;            // Show the value
                    series.DataLabels.ShowCategoryName = true;    // Show the category name
                }

                // Note: Auto‑scale is enabled by default; explicit property removed due to API change.

                // Determine output file path
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RadarChart_Output.xlsx");

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
