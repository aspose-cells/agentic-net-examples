// Title: How to verify a chart's secondary axis exists and safely assign series using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# check if a chart has a secondary axis before adding series | prevent runtime exception when setting IsOnSecondaryAxis in Aspose.Cells chart | add data series to the secondary axis of an Excel chart with Aspose.Cells .NET | use reflection to access IsOnSecondaryAxis property in Aspose.Cells C# | create secondary axis in an Aspose.Cells chart when it is missing
// Tags: Aspose.Cells chart secondary axis validation | C# set series on secondary axis Aspose.Cells | Aspose.Cells reflection IsOnSecondaryAxis handling | add dummy series to force secondary axis Aspose.Cells | prevent IsOnSecondaryAxis runtime error Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example loads an Excel workbook, ensures a chart is present, checks via reflection whether any series are already on a secondary axis, adds a dummy series to create the secondary axis if needed, then safely assigns the actual data series to the secondary axis, and finally saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure at least one chart exists; create a simple one if none are present
            if (worksheet.Charts.Count == 0)
            {
                int chartIdx = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
                Chart newChart = worksheet.Charts[chartIdx];
                newChart.NSeries.Add("A1:A5", true); // dummy series
            }

            // Access the first chart
            Chart chart = worksheet.Charts[0];

            // Determine whether a secondary axis already exists
            bool secondaryAxisExists = false;
            foreach (Series s in chart.NSeries)
            {
                // Use reflection to check for the property in case the API version differs
                var prop = s.GetType().GetProperty("IsOnSecondaryAxis");
                if (prop != null && prop.PropertyType == typeof(bool))
                {
                    if ((bool)prop.GetValue(s))
                    {
                        secondaryAxisExists = true;
                        break;
                    }
                }
            }

            // Helper to set the IsOnSecondaryAxis flag safely
            void SetSeriesOnSecondaryAxis(Series series)
            {
                var prop = series.GetType().GetProperty("IsOnSecondaryAxis");
                if (prop != null && prop.CanWrite)
                {
                    prop.SetValue(series, true);
                }
                else
                {
                    // If the property does not exist, ignore – the chart will use the primary axis.
                }
            }

            // If no secondary axis, add a dummy series to force its creation
            if (!secondaryAxisExists)
            {
                int dummyIdx = chart.NSeries.Add("A1", true);
                Series dummySeries = chart.NSeries[dummyIdx];
                SetSeriesOnSecondaryAxis(dummySeries);
            }

            // Add the actual series and assign it to the secondary axis
            int actualIdx = chart.NSeries.Add("B2:B5", true);
            Series actualSeries = chart.NSeries[actualIdx];
            SetSeriesOnSecondaryAxis(actualSeries);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
