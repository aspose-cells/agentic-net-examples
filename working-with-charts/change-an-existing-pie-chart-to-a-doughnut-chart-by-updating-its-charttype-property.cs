// Title: Aspose.Cells for .NET: Convert a Pie Chart to a Doughnut Chart (C#)
// Description: Load an Excel workbook with Aspose.Cells, locate the first chart, set its ChartType to Doughnut, and save the file. The sample includes checks for missing files and absent charts to prevent runtime errors.
// Keywords: Aspose.Cells C# change chart type | ChartType Doughnut | convert pie chart to doughnut Aspose | programmatic Excel chart update | C# Aspose.Cells chart manipulation | global | USA | India
// Common Searches: Aspose.Cells set chart type to Doughnut C# | how to replace a pie chart with a doughnut chart in Excel using .NET | change chart.Type to Doughnut Aspose.Cells example | C# code to convert existing pie chart to doughnut chart | update Excel chart programmatically Aspose.Cells
// Developer Intent: Modify the ChartType of an existing pie chart so it becomes a doughnut chart within an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Standardize report visuals by converting all pie charts to doughnut charts before distribution. | Automate chart‑type migration in a batch process that processes multiple workbooks. | Replace a legacy pie chart with a doughnut chart while preserving data series and formatting.
// AI Prompts: Generate C# code that loads a workbook, finds the first chart, changes its type to Doughnut, and saves the result using Aspose.Cells. | Write a method that scans every worksheet, detects pie charts, and switches each to a doughnut chart. | Explain best practices for handling workbooks that may lack charts when updating chart types with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace Example
{
    // Load an Excel workbook with Aspose.Cells, locate the first chart, set its ChartType to Doughnut, and save the file. The sample includes checks for missing files and absent charts to prevent runtime errors.
    class Program
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
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure there is at least one chart on the worksheet
                if (sheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the worksheet.");
                    return;
                }

                // Retrieve the first chart (assumed to be a pie chart)
                Chart chart = sheet.Charts[0];

                // Change the chart type to Doughnut
                chart.Type = ChartType.Doughnut;

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
