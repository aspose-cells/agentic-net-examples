// Title: Change a Pie Chart to a Doughnut Chart in an Existing Excel Workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code using Aspose.Cells that opens an existing .xlsx file, locates the first chart, sets its Type to ChartType.Doughnut, and saves the workbook. | Show how to programmatically replace a pie chart with a doughnut chart in a .NET application by modifying the Chart.Type property of an Aspose.Cells chart object.
// Common Searches: aspnet aspose.cells change chart type from pie to doughnut in existing workbook | c# programmatically convert Excel pie chart to doughnut using Aspose.Cells | how to update chart type of a chart already in an .xlsx file with Aspose.Cells .NET | Aspose.Cells modify first worksheet chart to doughnut chart C# example | replace pie chart with doughnut chart in Excel file using Aspose.Cells library
// Tags: Aspose.Cells chart type conversion | C# modify Excel chart to doughnut | Aspose.Cells load workbook edit chart | Chart.Type property doughnut Aspose.Cells | update existing Excel chart .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example loads an existing Excel workbook, checks for a chart on the first worksheet, changes the first chart's type from Pie to Doughnut via the Chart.Type property, and saves the modified file to a new location.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one chart on the sheet
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the worksheet.");
                    return;
                }

                // Retrieve the first chart (assumed to be a pie chart)
                Chart chart = worksheet.Charts[0];

                // Change the chart type from Pie to Doughnut
                chart.Type = ChartType.Doughnut;

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
}
