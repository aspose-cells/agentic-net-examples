// Title: Set PivotChart axis titles and enable drop zones with Aspose.Cells for .NET (C#)
// Description: Loads an existing XLSX workbook, finds the first PivotChart, assigns custom text to the Category and Value axes, activates all drop‑zone features via PivotOptions, and saves the updated file.
// Keywords: Aspose.Cells | C# | .NET | PivotChart | axis title | custom axis label | PivotOptions | DropZonesVisible | enable drop zones | Excel chart automation | programmatic chart customization | load workbook | save workbook | Excel automation
// Common Searches: Aspose.Cells set PivotChart axis title C# | Enable drop zones on PivotChart using Aspose.Cells | How to change PivotChart axis labels with .NET | PivotOptions DropZonesVisible example | Find and modify PivotChart in existing workbook Aspose
// Developer Intent: Programmatically add custom axis titles and turn on all drop‑zone areas for a PivotChart, then save the workbook.
// Use Cases: Standardize axis labels across corporate reporting workbooks before distribution. | Prepare Excel dashboards that let end‑users drag fields onto chart drop zones for interactive analysis. | Batch‑process multiple files to ensure PivotCharts have consistent titles and interactive zones.
// AI Prompts: Write C# code that opens an XLSX file with Aspose.Cells, locates the first PivotChart, sets custom Category and Value axis titles, enables all drop zones via PivotOptions, and saves the workbook. | Show robust error‑handling patterns for loading a workbook, detecting a PivotChart, and updating its PivotOptions in Aspose.Cells. | Explain how to verify that DropZonesVisible and related properties are applied after saving the workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an existing XLSX workbook, finds the first PivotChart, assigns custom text to the Category and Value axes, activates all drop‑zone features via PivotOptions, and saves the updated file.
class Program
{
    static void Main()
    {
        const string inputPath = "InputPivotChart.xlsx";
        const string outputPath = "OutputPivotChart.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook that contains a PivotChart
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Locate the first chart that is linked to a PivotTable
            Chart pivotChart = null;
            foreach (Chart chart in worksheet.Charts)
            {
                if (!string.IsNullOrEmpty(chart.PivotSource))
                {
                    pivotChart = chart;
                    break;
                }
            }

            if (pivotChart != null)
            {
                // Set custom titles for the chart axes
                pivotChart.CategoryAxis.Title.Text = "Custom Category Axis";
                pivotChart.ValueAxis.Title.Text = "Custom Value Axis";

                // Enable drop zones on the PivotChart
                PivotOptions options = pivotChart.PivotOptions;
                options.DropZonesVisible = true;
                options.DropZoneCategories = true;
                options.DropZoneData = true;
                options.DropZoneSeries = true;
                options.DropZoneFilter = true;
            }
            else
            {
                Console.WriteLine("No PivotChart found in the workbook.");
            }

            // Save the workbook with the updated chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
