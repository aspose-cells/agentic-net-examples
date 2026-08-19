// Title: Hide PivotChart Legend and Drop Zones with PivotOptions in Aspose.Cells for .NET (C#)
// Description: Loads a workbook, finds the first PivotChart, disables its legend (ShowLegend = false) and removes pivot drop zones (PivotOptions.DropZonesVisible = false), then saves the updated file.
// Keywords: Aspose.Cells C# hide PivotChart legend | PivotOptions DropZonesVisible false | remove pivot controls Aspose.Cells | chart formatting Aspose.Cells .NET | save workbook after chart changes
// Common Searches: how to hide legend of a PivotChart using Aspose.Cells | disable pivot drop zones on Excel chart C# | remove pivot controls from chart programmatically Aspose.Cells | save modified workbook after chart settings change
// Developer Intent: Programmatically hide a PivotChart legend and its pivot controls, then save the workbook.
// Use Cases: Produce clean reports by removing unnecessary legends from PivotCharts. | Lock chart interactivity by hiding pivot drop zones before distribution. | Batch‑process workbooks to enforce a consistent chart appearance across multiple files.
// AI Prompts: Write C# code that opens an Excel file with Aspose.Cells, hides legends and drop zones for every PivotChart, and saves the result. | Explain how PivotOptions.DropZonesVisible affects user interaction with a PivotChart in Aspose.Cells. | Create a loop that iterates through all worksheets and charts, applying legend and drop‑zone hiding to each PivotChart.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;

// Loads a workbook, finds the first PivotChart, disables its legend (ShowLegend = false) and removes pivot drop zones (PivotOptions.DropZonesVisible = false), then saves the updated file.
class Program
{
    static void Main()
    {
        const string inputPath = "Input.xlsx";
        const string outputPath = "Output.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Find the first chart that is linked to a PivotTable (PivotChart)
            Chart pivotChart = null;
            foreach (Chart ch in sheet.Charts)
            {
                if (!string.IsNullOrEmpty(ch.PivotSource))
                {
                    pivotChart = ch;
                    break;
                }
            }

            if (pivotChart != null)
            {
                // Hide the chart legend
                pivotChart.ShowLegend = false;

                // Hide all pivot controls on the chart
                PivotOptions pivotOptions = pivotChart.PivotOptions;
                if (pivotOptions != null)
                {
                    pivotOptions.DropZonesVisible = false;
                }
            }
            else
            {
                Console.WriteLine("No PivotChart found in the first worksheet.");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
