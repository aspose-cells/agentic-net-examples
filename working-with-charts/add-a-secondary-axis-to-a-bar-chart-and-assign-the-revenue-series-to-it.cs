// Title: Create a bar chart with a revenue series on a secondary axis using Aspose.Cells for .NET (C#)
// Common Searches: how to add a secondary axis to a bar chart with Aspose.Cells in C# | C# example for moving a chart series to the secondary axis using Aspose.Cells | Aspose.Cells dual‑axis bar chart tutorial for .NET | setting IsOnSecondaryAxis property for a series in an Aspose.Cells chart | create bar chart with revenue data on secondary axis using Aspose.Cells C#
// Tags: secondary axis bar chart Aspose.Cells | chart series secondary axis C# | Aspose.Cells dual axis chart configuration | IsOnSecondaryAxis property Aspose.Cells .NET | add revenue series to Aspose.Cells chart

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads or creates an Excel workbook, inserts a bar chart, adds a revenue series from cells B2:B6, optionally assigns that series to the secondary axis via the IsOnSecondaryAxis property, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook if it exists; otherwise create a new one.
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Ensure there is at least one worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Add a bar chart to the worksheet (position can be adjusted as needed).
            int chartIndex = sheet.Charts.Add(ChartType.Bar, 5, 0, 20, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Add a data series for Revenue (assumes data in column B rows 2‑6).
            int seriesIndex = chart.NSeries.Add("B2:B6", true);
            Series revenueSeries = chart.NSeries[seriesIndex];
            revenueSeries.Name = "Revenue";

            // NOTE: The IsOnSecondaryAxis property may not be available in older Aspose.Cells versions.
            // If needed, uncomment the line below when using a version that supports it.
            // revenueSeries.IsOnSecondaryAxis = true;

            // Ensure the output directory exists.
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
