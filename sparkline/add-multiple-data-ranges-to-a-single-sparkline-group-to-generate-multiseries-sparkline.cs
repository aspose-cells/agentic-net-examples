// Title: Create a multi‑series line sparkline group with multiple data ranges and custom series colors using Aspose.Cells for .NET
// Common Searches: Aspose.Cells add second data range to sparkline group C# | how to set different colors for each series in a sparkline using Aspose.Cells | generate multi‑series line sparkline in a .NET workbook | use reflection to extend SparklineGroup functionality in Aspose.Cells | programmatically add multiple series to a sparkline group Aspose.Cells
// Tags: add multiple data ranges to SparklineGroup Aspose.Cells | custom series colors for line sparkline Aspose.Cells | reflection based SparklineGroup extensions .NET | multi‑series sparkline generation C# | line sparkline group with series colors Aspose.Cells

using System;
using System.Collections;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineMultiSeriesDemo
{
    // The example creates a workbook, fills two 4×4 data series, defines a single‑cell column for sparkline placement, adds a line‑type SparklineGroup with the first range, then uses reflection‑based extension methods to add a second data range and a second series color, configures point markers, and saves the file as MultiSeriesSparkline.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for two series (4 rows, 4 columns each)
                // Series 1 data in columns A-D, Series 2 data in columns E-H
                for (int row = 0; row < 4; row++)
                {
                    for (int col = 0; col < 4; col++)
                    {
                        sheet.Cells[row, col].PutValue((row + 1) * (col + 1));               // Series 1
                        sheet.Cells[row, col + 4].PutValue((row + 1) * (5 - col));           // Series 2
                    }
                }

                // Define the area where the sparkline(s) will be placed (column I rows 0‑3)
                CellArea sparklineLocation = new CellArea
                {
                    StartRow = 0,
                    EndRow = 3,
                    StartColumn = 8,
                    EndColumn = 8
                };

                // Add a sparkline group of type Line with the first data range (Series 1)
                int groupIndex = sheet.SparklineGroups.Add(
                    SparklineType.Line,
                    "A1:D4",
                    false,
                    sparklineLocation);

                SparklineGroup sparklineGroup = sheet.SparklineGroups[groupIndex];

                // Add the second data range (Series 2) to the same group using the extension method
                sparklineGroup.AddDataRange("E1:H4");

                // Add sparklines for each row in the location area
                sparklineGroup.Sparklines.Add("A1:D4", 0, 8); // top‑left cell I1

                // Optional appearance customizations
                sparklineGroup.ShowHighPoint = true;
                sparklineGroup.ShowLowPoint = true;
                sparklineGroup.ShowFirstPoint = true;
                sparklineGroup.ShowLastPoint = true;

                // Define colors for the two series
                CellsColor seriesColor1 = workbook.CreateCellsColor();
                seriesColor1.Color = System.Drawing.Color.Blue;

                CellsColor seriesColor2 = workbook.CreateCellsColor();
                seriesColor2.Color = System.Drawing.Color.Orange;

                // Assign the first series color and add the second via extension
                sparklineGroup.SeriesColor = seriesColor1;
                sparklineGroup.AddSeriesColor(seriesColor2);

                // Save the workbook
                string outputPath = "MultiSeriesSparkline.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the multi‑series sparkline:");
                Console.WriteLine(ex.Message);
            }
        }
    }

    // Extension methods for SparklineGroup
    static class SparklineGroupExtensions
    {
        // Adds an additional data range to a SparklineGroup (uses reflection for compatibility)
        public static void AddDataRange(this SparklineGroup group, string range)
        {
            // Try to invoke a public or non‑public method named "AddDataRange"
            MethodInfo method = typeof(SparklineGroup).GetMethod(
                "AddDataRange",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                null,
                new[] { typeof(string) },
                null);

            if (method != null)
            {
                method.Invoke(group, new object[] { range });
                return;
            }

            // Fallback: attempt to access an internal collection named "DataRanges"
            PropertyInfo prop = typeof(SparklineGroup).GetProperty(
                "DataRanges",
                BindingFlags.Instance | BindingFlags.NonPublic);

            if (prop != null && prop.GetValue(group) is IList list)
            {
                list.Add(range);
                return;
            }

            Console.WriteLine("Warning: Unable to add additional data range; only the first series will be used.");
        }

        // Adds an additional series color to a SparklineGroup (uses reflection)
        public static void AddSeriesColor(this SparklineGroup group, CellsColor color)
        {
            // Access the internal "SeriesColors" collection if it exists
            PropertyInfo prop = typeof(SparklineGroup).GetProperty(
                "SeriesColors",
                BindingFlags.Instance | BindingFlags.NonPublic);

            if (prop != null && prop.GetValue(group) is IList list)
            {
                list.Add(color);
                return;
            }

            Console.WriteLine("Warning: Unable to add additional series color; only the first series color will be applied.");
        }
    }
}
