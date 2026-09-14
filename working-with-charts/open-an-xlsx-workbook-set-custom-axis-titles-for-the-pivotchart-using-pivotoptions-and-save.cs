// Title: Set custom X‑axis and Y‑axis titles on a PivotChart in an existing XLSX workbook using Aspose.Cells for .NET
// AI Prompts: Open an existing XLSX file with Aspose.Cells, locate the first PivotChart, enable axis title visibility, assign new text to the category and value axes, and save the workbook. | Programmatically change the category (X) axis label and the value (Y) axis label of a PivotChart in C# using Aspose.Cells' Axis.Title properties. | Update a workbook's PivotChart to display custom axis titles and write the modified file to a new location with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# set pivot chart X axis title in existing workbook | how to change Y axis label of a PivotChart using Aspose.Cells for .NET | modify axis titles of a chart in an XLSX file with Aspose.Cells C# example | update pivot chart axis text programmatically Aspose.Cells
// Tags: pivot chart axis title Aspose.Cells | set category axis text .NET | modify value axis label C# | Aspose.Cells update chart axis labels XLSX | axis title modification workbook

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example loads an existing XLSX workbook, verifies a chart exists, assumes the first chart is a PivotChart, makes the category and value axis titles visible, sets custom text for each axis, and saves the updated workbook to a new file, handling missing files or axes gracefully.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing XLSX workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (adjust index if needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure there is at least one chart on the sheet
                if (sheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the worksheet.");
                    return;
                }

                // Assume the first chart on the sheet is a PivotChart
                Chart pivotChart = sheet.Charts[0];

                // Set custom titles for the category (X) and value (Y) axes
                // Ensure the axis titles are visible
                if (pivotChart.CategoryAxis != null && pivotChart.ValueAxis != null)
                {
                    // Category axis (usually X axis)
                    Axis categoryAxis = pivotChart.CategoryAxis;
                    categoryAxis.Title.IsVisible = true;
                    categoryAxis.Title.Text = "Custom Category Axis Title";

                    // Value axis (usually Y axis)
                    Axis valueAxis = pivotChart.ValueAxis;
                    valueAxis.Title.IsVisible = true;
                    valueAxis.Title.Text = "Custom Value Axis Title";
                }
                else
                {
                    Console.WriteLine("The chart does not have both category and value axes to set titles.");
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
