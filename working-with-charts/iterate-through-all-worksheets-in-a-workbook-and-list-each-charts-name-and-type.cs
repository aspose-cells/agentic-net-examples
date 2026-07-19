// Title: C# – List All Chart Names and Types Across Worksheets with Aspose.Cells
// Description: Loads a workbook, iterates through every worksheet, accesses each sheet's ChartCollection, and prints the chart index, name (or "(unnamed)"), and ChartType to the console using Aspose.Cells for .NET.
// Keywords: Aspose.Cells chart enumeration | C# list chart names | retrieve chart type Aspose.Cells | iterate worksheets charts | Aspose.Cells ChartCollection example
// Common Searches: how to get chart names from all sheets using Aspose.Cells | C# loop through worksheets and display chart type | Aspose.Cells enumerate charts per worksheet | list chart properties in a workbook .NET | Aspose.Cells get chart index and type
// Developer Intent: Extract and display the name and type of every chart present in each worksheet of a workbook.
// Use Cases: Generate an audit report of chart assets for compliance checks. | Validate that required chart types exist before exporting or publishing a workbook. | Log chart metadata during automated UI testing to ensure correct visualization generation.
// AI Prompts: Create C# code that writes each chart's name and type from all worksheets into a CSV file using Aspose.Cells. | Show how to assign default names to unnamed charts while iterating through a workbook's sheets. | Provide an example that filters only pie charts across every worksheet and outputs their details.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartInfo
{
    // Loads a workbook, iterates through every worksheet, accesses each sheet's ChartCollection, and prints the chart index, name (or "(unnamed)"), and ChartType to the console using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the chart collection of the current worksheet
                ChartCollection charts = sheet.Charts;

                // If there are no charts, continue to the next worksheet
                if (charts.Count == 0)
                {
                    continue;
                }

                // List each chart's name and type
                for (int i = 0; i < charts.Count; i++)
                {
                    Chart chart = charts[i];
                    string chartName = string.IsNullOrEmpty(chart.Name) ? "(unnamed)" : chart.Name;
                    ChartType chartType = chart.Type;

                    Console.WriteLine($"Worksheet: {sheet.Name}, Chart Index: {i}, Name: {chartName}, Type: {chartType}");
                }
            }

            // Optionally, save the workbook after any modifications (if needed)
            // workbook.Save("output.xlsx");
        }
    }
}
