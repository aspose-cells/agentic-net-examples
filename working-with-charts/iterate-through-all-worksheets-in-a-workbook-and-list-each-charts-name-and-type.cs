// Title: Enumerate chart names and types on every worksheet of an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write a C# console program that loads an .xlsx file with Aspose.Cells, loops through all worksheets, and prints each chart’s Name and Type. | Show how to access the ChartCollection of each Worksheet in Aspose.Cells and output the chart properties to the console. | Provide sample code that saves the workbook after enumerating charts without modifying the file, using Aspose.Cells for .NET.
// Common Searches: asp.net aspose.cells list chart name and chart type for each sheet | c# iterate all worksheets and get chart information using Aspose.Cells | how to retrieve chart type from worksheet charts in Aspose.Cells .NET | example code to print chart details from an Excel workbook with Aspose.Cells
// Tags: enumerate worksheet chart properties Aspose.Cells C# | retrieve chart name and type Aspose.Cells .NET | loop through chart collection per worksheet Aspose.Cells | output Excel chart metadata console Aspose.Cells | read-only chart enumeration Aspose.Cells workbook

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartInfo
{
    // The sample loads 'input.xlsx', iterates over every worksheet, accesses each sheet's ChartCollection, and writes the worksheet name, chart name, and chart type to the console before optionally saving the workbook as 'output.xlsx'.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (load rule)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the chart collection of the current worksheet
                ChartCollection charts = sheet.Charts;

                // Iterate through each chart in the collection
                for (int i = 0; i < charts.Count; i++)
                {
                    Chart chart = charts[i];

                    // Output worksheet name, chart name and chart type
                    Console.WriteLine($"Worksheet: {sheet.Name}, Chart Name: {chart.Name}, Chart Type: {chart.Type}");
                }
            }

            // Save the workbook (save rule) – optional if no changes were made
            workbook.Save("output.xlsx");
        }
    }
}
