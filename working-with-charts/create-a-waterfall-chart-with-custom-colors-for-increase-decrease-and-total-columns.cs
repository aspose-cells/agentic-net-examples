// Title: Generate a Waterfall chart in Excel with custom increase, decrease, and total column colors using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates an Excel workbook, adds a Waterfall chart, and assigns green to positive columns, red to negative columns, and blue to the total column with Aspose.Cells. | Update an existing Aspose.Cells Waterfall chart so each point’s foreground color is set conditionally based on its value and whether it represents the total.
// Common Searches: how to apply different colors to positive and negative points in an Aspose.Cells waterfall chart c# | Aspose.Cells C# set total column color in waterfall chart | customize point colors in Excel waterfall chart using Aspose.Cells .NET | example of coloring increase decrease columns in Aspose.Cells waterfall chart
// Tags: Aspose.Cells waterfall chart point coloring | C# set chart series point foreground color | Excel waterfall chart custom increase decrease colors | Aspose.Cells conditional point formatting | generate colored waterfall chart .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates an Excel workbook, fills it with category and value data, adds a Waterfall chart, and applies green for increase points, red for decrease points, and blue for the final total point before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet's cells
            var workbook = new Workbook();
            var cells = workbook.Worksheets[0].Cells;

            // Sample data for the waterfall chart
            // Column A: Category names
            // Column B: Values (positive for increase, negative for decrease)
            string[] categories = { "Start", "Revenue", "Cost", "Profit", "Tax", "Net Income" };
            double[] values = { 100, 150, -50, 100, -30, 70 };

            // Populate worksheet with data
            for (int i = 0; i < categories.Length; i++)
            {
                cells[i, 0].PutValue(categories[i]); // Category
                cells[i, 1].PutValue(values[i]);    // Value
            }

            // Add a Waterfall chart to the worksheet
            int chartIndex = workbook.Worksheets[0].Charts.Add(ChartType.Waterfall, 5, 0, 25, 10);
            var chart = workbook.Worksheets[0].Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B1:B6", true);
            chart.NSeries.CategoryData = "A1:A6";

            // Define custom colors
            Color increaseColor = Color.Green;   // Positive (increase) columns
            Color decreaseColor = Color.Red;     // Negative (decrease) columns
            Color totalColor = Color.Blue;       // Total column (last point)

            // Apply custom colors to each point based on its value
            for (int i = 0; i < chart.NSeries[0].Points.Count; i++)
            {
                var point = chart.NSeries[0].Points[i];
                double val = values[i];

                if (i == categories.Length - 1) // Last point is the total column
                {
                    point.Area.ForegroundColor = totalColor;
                }
                else if (val >= 0) // Increase
                {
                    point.Area.ForegroundColor = increaseColor;
                }
                else // Decrease
                {
                    point.Area.ForegroundColor = decreaseColor;
                }
            }

            // Determine output path and ensure directory exists
            string outputFile = "WaterfallChart.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputFile)) ?? string.Empty;
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the chart
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
