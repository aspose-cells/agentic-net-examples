// Title: Aspose.Cells C# – Waterfall Chart with Custom Start, Intermediate & Total Colors
// Description: Creates a workbook, adds sample data, inserts a Waterfall chart, and programmatically colors the start point (green), intermediate points (light blue) and total point (orange) using Aspose.Cells for .NET before saving the file.
// Keywords: Aspose.Cells waterfall chart C# | custom point colors Aspose.Cells | set start point color waterfall chart | intermediate point color Aspose.Cells | total point color waterfall chart .NET | C# chart point formatting Aspose | Excel waterfall chart custom colors
// Common Searches: how to change waterfall chart point colors in Aspose.Cells | Aspose.Cells set start and total colors for waterfall chart | C# example for custom colored waterfall chart | waterfall chart color coding Aspose.Cells .NET | highlight start and total points in Excel waterfall chart programmatically
// Developer Intent: Generate a waterfall chart and apply distinct fill colors to the start, intermediate, and total data points using Aspose.Cells for .NET.
// Use Cases: Financial statements where opening balance, quarterly changes, and closing balance need visual distinction. | Project budgeting reports that highlight initial budget, incremental expenses, and final total. | Sales funnel analysis with separate colors for each stage and the overall result.
// AI Prompts: Write C# code with Aspose.Cells to color the first and last points of a waterfall chart differently from the middle points. | Explain how to detect start, intermediate, and total points in a waterfall series and assign specific RGB colors. | Show how to add a legend entry for each point type while using custom colors in an Aspose.Cells waterfall chart.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace WaterfallChartDemo
{
    // Creates a workbook, adds sample data, inserts a Waterfall chart, and programmatically colors the start point (green), intermediate points (light blue) and total point (orange) using Aspose.Cells for .NET before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a waterfall chart
                // Column A – Categories, Column B – Values
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");

                sheet.Cells["A2"].PutValue("Start");
                sheet.Cells["B2"].PutValue(5000);   // Start point

                sheet.Cells["A3"].PutValue("Q1");
                sheet.Cells["B3"].PutValue(3000);   // Intermediate point

                sheet.Cells["A4"].PutValue("Q2");
                sheet.Cells["B4"].PutValue(-2000);  // Intermediate point

                sheet.Cells["A5"].PutValue("Q3");
                sheet.Cells["B5"].PutValue(4000);   // Intermediate point

                sheet.Cells["A6"].PutValue("Total");
                sheet.Cells["B6"].PutValue(0);      // Total point (value will be calculated by Excel)

                // Add a waterfall chart
                int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 8, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series (values) and categories
                chart.NSeries.Add("B2:B6", true);
                chart.NSeries.CategoryData = "A2:A6";

                // Apply distinct colors to points
                for (int i = 0; i < chart.NSeries[0].Points.Count; i++)
                {
                    ChartPoint point = chart.NSeries[0].Points[i];

                    if (i == 0) // Start point
                    {
                        point.Area.ForegroundColor = Color.Green;
                    }
                    else if (i == chart.NSeries[0].Points.Count - 1) // Total point
                    {
                        // Excel treats the last point as total automatically for waterfall charts
                        point.Area.ForegroundColor = Color.Orange;
                    }
                    else // Intermediate points
                    {
                        point.Area.ForegroundColor = Color.LightBlue;
                    }
                }

                // Define output file path
                string outputPath = "WaterfallChartDemo.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
