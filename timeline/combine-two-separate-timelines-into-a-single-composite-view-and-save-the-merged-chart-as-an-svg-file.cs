// Title: Create a dual‑axis timeline line chart and export it as an SVG image with Aspose.Cells for .NET (C#)
// Common Searches: how to create a line chart with two timelines on primary and secondary axes using Aspose.Cells in C# | export an Excel chart to SVG format with Aspose.Cells for .NET | combine multiple series into a single chart and render it as SVG in C# | set a secondary axis for a series in an Aspose.Cells line chart | save a chart as an SVG file using Aspose.Cells ImageOrPrintOptions C#
// Tags: Aspose.Cells dual axis line chart | export chart to SVG Aspose.Cells | combined timeline chart C# | render Excel chart as SVG using ImageOrPrintOptions | secondary axis series Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace MergeTimelinesToSvg
{
    // The example creates a new workbook, fills columns with dates and two sets of values, adds a line chart that plots both series (placing the second series on a secondary axis), sets chart titles and legend, and then renders the chart to an SVG file using ImageOrPrintOptions in Aspose.Cells for .NET (C#).
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Populate sample data for two separate timelines
                // Header row
                sheet.Cells["A1"].PutValue("Date");
                sheet.Cells["B1"].PutValue("Timeline 1");
                sheet.Cells["C1"].PutValue("Timeline 2");

                // Sample dates and values
                DateTime startDate = new DateTime(2023, 1, 1);
                for (int i = 0; i < 6; i++)
                {
                    // Date column (store as DateTime)
                    sheet.Cells[i + 1, 0].PutValue(startDate.AddMonths(i));

                    // Timeline 1 values
                    sheet.Cells[i + 1, 1].PutValue(100 + i * 20);

                    // Timeline 2 values (different scale)
                    sheet.Cells[i + 1, 2].PutValue(50 + i * 15);
                }

                // Format the date column for readability
                Style dateStyle = workbook.CreateStyle();
                dateStyle.Number = 14; // "m/d/yyyy"
                sheet.Cells.Columns[0].ApplyStyle(dateStyle, new StyleFlag { NumberFormat = true });

                // 3. Add a line chart that combines both timelines
                int chartIndex = sheet.Charts.Add(ChartType.Line, 8, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Add first series (Timeline 1) – primary axis
                chart.NSeries.Add("B2:B7", true);
                chart.NSeries[0].Name = "Timeline 1";

                // Add second series (Timeline 2) – secondary axis
                chart.NSeries.Add("C2:C7", true);
                chart.NSeries[1].Name = "Timeline 2";
                chart.NSeries[1].PlotOnSecondAxis = true; // place on secondary axis

                // Set the category (X) axis data – dates from column A
                chart.NSeries.CategoryData = "A2:A7";

                // Chart titles
                chart.Title.Text = "Combined Timelines";
                chart.Title.Font.Size = 12;
                chart.CategoryAxis.Title.Text = "Date";
                chart.ValueAxis.Title.Text = "Timeline 1 Values";

                // Optional: improve appearance
                chart.Legend.Position = LegendPositionType.Bottom;

                // 4. Render the combined chart to an SVG file
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = Aspose.Cells.Drawing.ImageType.Svg,
                    OnePagePerSheet = true
                };

                string outputPath = Path.Combine(Environment.CurrentDirectory, "CombinedTimelines.svg");

                // Save the chart as SVG
                chart.ToImage(outputPath, imgOptions);

                Console.WriteLine("Combined timeline chart successfully saved as SVG at:");
                Console.WriteLine(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the combined timeline SVG:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
