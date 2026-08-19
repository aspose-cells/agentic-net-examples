// Title: List All Chart Names and Types Across Worksheets with Aspose.Cells (C#)
// Description: C# example that creates a workbook, adds sample charts, then iterates every worksheet to output each chart’s name (or a placeholder for unnamed charts) and its ChartType. The workbook is saved after the enumeration.
// Keywords: Aspose.Cells | C# | enumerate charts | chart name | chart type | list charts workbook | iterate worksheets | chart metadata | Aspose.Cells chart enumeration | retrieve chart information
// Common Searches: Aspose.Cells get chart name and type | C# list charts in workbook | How to enumerate charts in Aspose.Cells | Retrieve chart metadata from each worksheet Aspose.Cells | Aspose.Cells iterate worksheets and charts
// Developer Intent: Retrieve each chart’s name and its ChartType from every worksheet in a workbook.
// Use Cases: Generate a quick inventory of all charts for documentation or review. | Validate that required chart types are present before exporting or publishing a workbook. | Log chart metadata for auditing, debugging, or automated quality checks.
// AI Prompts: Write C# code using Aspose.Cells that enumerates all worksheets and writes each chart’s name and type to a CSV file. | Show how to filter only Pie charts while iterating through charts across worksheets with Aspose.Cells. | Explain safe handling of charts without a name when listing chart information in Aspose.Cells. | Create a PowerShell script that calls a compiled .NET assembly to list chart names and types from an Excel file.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartInfo
{
    // C# example that creates a workbook, adds sample charts, then iterates every worksheet to output each chart’s name (or a placeholder for unnamed charts) and its ChartType. The workbook is saved after the enumeration.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook with a default worksheet
                Workbook workbook = new Workbook();

                // Ensure we have two worksheets for the demo
                if (workbook.Worksheets.Count < 2)
                {
                    workbook.Worksheets.Add(); // adds a second worksheet
                }

                // -------------------------------------------------
                // Add sample data and charts to demonstrate the logic
                // -------------------------------------------------
                for (int wsIndex = 0; wsIndex < 2; wsIndex++)
                {
                    Worksheet sheet = workbook.Worksheets[wsIndex];
                    sheet.Name = $"Sheet{wsIndex + 1}";

                    // Populate some data
                    sheet.Cells["A1"].PutValue("Category");
                    sheet.Cells["B1"].PutValue("Value");
                    sheet.Cells["A2"].PutValue("A");
                    sheet.Cells["A3"].PutValue("B");
                    sheet.Cells["A4"].PutValue("C");
                    sheet.Cells["B2"].PutValue(10 + wsIndex * 5);
                    sheet.Cells["B3"].PutValue(20 + wsIndex * 5);
                    sheet.Cells["B4"].PutValue(30 + wsIndex * 5);

                    // Add a chart of different type per worksheet
                    ChartType type = wsIndex == 0 ? ChartType.Column : ChartType.Pie;
                    int chartIdx = sheet.Charts.Add(type, 5, 0, 15, 8);
                    Chart chart = sheet.Charts[chartIdx];
                    chart.NSeries.Add("B2:B4", true);
                    chart.NSeries.CategoryData = "A2:A4";

                    // Optionally give the chart a name
                    chart.Name = $"Chart_{wsIndex + 1}";
                }

                // -------------------------------------------------
                // Iterate through all worksheets and list each chart's name and type
                // -------------------------------------------------
                Console.WriteLine("Charts in the workbook:");
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    if (ws.Charts.Count == 0)
                        continue;

                    Console.WriteLine($"Worksheet: {ws.Name}");
                    for (int i = 0; i < ws.Charts.Count; i++)
                    {
                        Chart chart = ws.Charts[i];
                        string chartName = string.IsNullOrEmpty(chart.Name) ? "(unnamed)" : chart.Name;
                        Console.WriteLine($"  Chart {i + 1}: Name = {chartName}, Type = {chart.Type}");
                    }
                }

                // Save the workbook (optional, just to demonstrate lifecycle usage)
                workbook.Save("ChartsInfoDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
