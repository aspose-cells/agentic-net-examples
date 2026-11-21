// Title: How to enable cell‑based data labels for a chart series using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set chart series data labels to use cell range | Enable cell‑based data labels on an Excel chart using Aspose.Cells .NET | Show data labels from worksheet cells in a chart series with Aspose.Cells | C# Aspose.Cells example for Series.DataLabels.ShowCellRange property | Programmatically activate cell range data labels in Excel chart via Aspose.Cells
// Tags: Aspose.Cells chart series ShowCellRange | Series.DataLabels.ShowCellRange usage | Excel chart cell range labels .NET | Aspose.Cells workbook chart manipulation C# | set chart series data labels programmatically

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads an existing Excel workbook, checks for a chart and its first series, sets Series.DataLabels.ShowCellRange to true to display data labels sourced from worksheet cells, and saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one chart on the worksheet
            if (sheet.Charts.Count > 0)
            {
                // Get the first chart
                Chart chart = sheet.Charts[0];

                // Ensure the chart has at least one series
                if (chart.NSeries.Count > 0)
                {
                    // Get the first series
                    Series series = chart.NSeries[0];

                    // Activate cell‑based data labels for this series
                    series.DataLabels.ShowCellRange = true;
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
