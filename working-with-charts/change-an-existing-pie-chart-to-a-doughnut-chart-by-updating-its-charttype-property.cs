// Title: Change an existing Excel pie chart to a doughnut chart using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells change chart type from pie to doughnut | How to convert a pie chart to a doughnut chart in an existing workbook with Aspose.Cells | Set Chart.Type to Doughnut using Aspose.Cells .NET example | Programmatically update Excel chart to doughnut with Aspose.Cells API | Modify chart type in Aspose.Cells C# code
// Tags: Aspose.Cells chart type conversion | C# modify Excel chart to doughnut | Aspose.Cells Chart.Type property usage | Excel workbook chart update Aspose.Cells | Doughnut chart generation with Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example loads an existing Excel file, accesses the first worksheet's chart, changes its Type property to ChartType.Doughnut, and saves the workbook, demonstrating how to convert a pie chart to a doughnut chart with Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook containing the chart
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one chart
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the worksheet.");
                    return;
                }

                // Get the first chart on the worksheet
                Chart chart = worksheet.Charts[0];

                // Change the chart type from Pie to Doughnut
                chart.Type = ChartType.Doughnut;

                // Save the workbook with the updated chart
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
