// Title: Detect if an Excel chart X‑axis is a value axis and convert it to a scatter chart using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells how to determine whether a chart X axis is a value axis | convert Excel chart X axis to numeric axis with Aspose.Cells .NET | change chart type to scatter programmatically using Aspose.Cells C# | set XValues range for scatter chart series in Aspose.Cells | Aspose.Cells check and modify chart axis type before saving workbook
// Tags: chart X axis value detection Aspose.Cells | convert chart to scatter Aspose.Cells C# | assign series XValues range Aspose.Cells | update chart axis titles Aspose.Cells | save workbook after chart modification Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an Excel workbook, verifies the presence of a chart, checks if the chart is a Scatter type (which uses a numeric X‑axis), and if not, changes the chart type to Scatter, assigns explicit XValues and Y values ranges for each series, updates axis titles, and saves the modified workbook.
class ChartAxisConverter
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart
            Chart chart = sheet.Charts[0];

            // Determine if the chart already uses a value axis for X (i.e., a scatter chart)
            bool isXAxisValueAxis = chart.Type == ChartType.Scatter;

            if (isXAxisValueAxis)
            {
                Console.WriteLine("The X axis is already a Value axis (scatter chart).");
            }
            else
            {
                Console.WriteLine("X axis is not a Value axis. Converting...");

                // Convert the chart to a basic Scatter chart (uses value axis for X)
                chart.Type = ChartType.Scatter;

                // For a Scatter chart, each series must have explicit X values.
                // Adjust the ranges according to your data layout.
                foreach (Series series in chart.NSeries)
                {
                    // Example: set X values to column A and Y values to column B.
                    series.XValues = "Sheet1!A2:A6";
                    series.Values = "Sheet1!B2:B6";
                }

                // Optionally, adjust axis titles to reflect numeric X axis.
                if (chart.CategoryAxis?.Title != null)
                {
                    chart.CategoryAxis.Title.Text = "Numeric X Axis";
                }

                if (chart.ValueAxis?.Title != null)
                {
                    chart.ValueAxis.Title.Text = "Y Axis";
                }

                Console.WriteLine("Conversion completed.");
            }

            // Ensure output directory exists
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
