// Title: Save Modified Chart Workbook to XLSX Without Losing Controls – Aspose.Cells C#
// Description: Loads an existing XLSX file, checks for charts, clears and adds a data series, updates the chart title, and saves the workbook as XLSX while preserving all chart objects and interactive controls using Aspose.Cells for .NET.
// Keywords: Aspose.Cells save chart | preserve chart controls XLSX | modify chart series C# | update chart title Aspose.Cells | load workbook edit chart | Aspose.Cells chart serialization | C# Excel chart manipulation
// Common Searches: how to save an Aspose.Cells workbook with edited charts | preserve chart objects when exporting to XLSX using Aspose.Cells | update chart series and title then save workbook .NET | Aspose.Cells keep legends and axes after saving | C# code to modify Excel chart and retain controls
// Developer Intent: Update a chart in an existing workbook and save the file as XLSX without losing any chart controls.
// Use Cases: Programmatically replace the data range of the first chart, change its title, and export the workbook while keeping legends, axes, and interactivity. | Validate that a worksheet contains at least one chart before applying modifications to avoid runtime errors. | Automate batch processing of Excel files where chart layouts must remain unchanged after data updates.
// AI Prompts: Generate C# code that loads an XLSX workbook, modifies the first chart's series and title, and saves it as XLSX while preserving all chart controls with Aspose.Cells. | Explain the mechanisms Aspose.Cells uses to retain chart objects during SaveFormat.Xlsx and list any options that affect chart serialization. | Create robust error‑handling for checking file existence, worksheet presence, and chart count before updating and saving a workbook with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartSaveExample
{
    // Loads an existing XLSX file, checks for charts, clears and adds a data series, updates the chart title, and saves the workbook as XLSX while preserving all chart objects and interactive controls using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found. Please provide a valid workbook.");
                    return;
                }

                // Load the existing workbook that contains a chart
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure the worksheet has at least one chart
                if (sheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the worksheet.");
                    return;
                }

                // Get the first chart
                Chart chart = sheet.Charts[0];

                // Modify the chart data: clear existing series and add a new series
                chart.NSeries.Clear();
                chart.NSeries.Add("B2:B5", true);
                // Category data line removed because the Series class does not expose a CategoryData property in this version

                // Optionally, modify chart title or other properties
                chart.Title.Text = "Updated Chart Title";

                // Save the workbook to XLSX format, preserving all chart controls
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\" with updated chart.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
