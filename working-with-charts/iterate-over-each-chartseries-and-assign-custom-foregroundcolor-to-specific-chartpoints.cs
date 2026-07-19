// Title: Set individual ChartPoint foreground colors in all ChartSeries of a column chart using Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds category data and two numeric series, inserts a column chart, then loops through each ChartSeries and its ChartPoints to assign a custom Area.ForegroundColor from a predefined palette and applies FormattingType.Custom before saving the file as an XLSX workbook.
// Keywords: Aspose.Cells chart point color C# | set foreground color ChartPoint Aspose.Cells | customize column chart series Aspose.Cells .NET | iterate chart series points Aspose.Cells | FormattingType.Custom chart Aspose | apply colors to chart points .NET | Aspose.Cells chart styling example | C# Excel chart point formatting | Aspose.Cells multi‑color column chart
// Common Searches: how to change color of individual columns in Aspose.Cells chart C# | Aspose.Cells set foreground color for each chart point | iterate over chart series and points Aspose.Cells .NET | apply custom formatting to chart points Aspose.Cells | C# example for coloring chart points in Excel with Aspose
// Developer Intent: The developer wants to programmatically assign distinct foreground colors to every data point across all series in a column chart using Aspose.Cells for .NET.
// Use Cases: Highlight outlier values by coloring specific columns differently. | Create a multi‑colored column chart that uses a consistent palette for each category. | Implement conditional formatting on chart points to visualize value ranges without using Excel's built‑in rules.
// AI Prompts: Write C# code with Aspose.Cells that loops through each ChartSeries in a column chart and sets a custom ForegroundColor for every ChartPoint using a Color array. | Show how to apply FormattingType.Custom to chart points after assigning their Area.ForegroundColor in Aspose.Cells for .NET. | Provide a complete example that creates sample data, adds a column chart, colors each point individually, and saves the workbook as an XLSX file.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds category data and two numeric series, inserts a column chart, then loops through each ChartSeries and its ChartPoints to assign a custom Area.ForegroundColor from a predefined palette and applies FormattingType.Custom before saving the file as an XLSX workbook.
    public class ChartSeriesPointsForegroundColor
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["A5"].PutValue("D");

                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["B5"].PutValue(40);

                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["C2"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["C4"].PutValue(35);
                sheet.Cells["C5"].PutValue(45);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series (both series)
                chart.NSeries.Add("B2:C5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Define custom colors for demonstration
                Color[] pointColors = new Color[]
                {
                    Color.FromArgb(255, 99, 71),   // Tomato
                    Color.FromArgb(60, 179, 113), // MediumSeaGreen
                    Color.FromArgb(30, 144, 255), // DodgerBlue
                    Color.FromArgb(218, 112, 214) // Orchid
                };

                // Apply foreground colors to each point in each series
                for (int s = 0; s < chart.NSeries.Count; s++)
                {
                    Series series = chart.NSeries[s];
                    ChartPointCollection points = series.Points;

                    for (int p = 0; p < points.Count; p++)
                    {
                        ChartPoint point = points[p];
                        point.Area.ForegroundColor = pointColors[p % pointColors.Length];
                        point.Area.Formatting = FormattingType.Custom; // Ensure custom formatting is applied
                    }
                }

                // Save the workbook to a file
                string outputPath = "ChartSeriesPointsForegroundColor.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ChartSeriesPointsForegroundColor.Run();
        }
    }
}
