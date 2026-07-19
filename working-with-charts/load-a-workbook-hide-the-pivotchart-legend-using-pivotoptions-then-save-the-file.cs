// Title: Hide PivotChart Legend and Drop Zones with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, verifies a chart on the first worksheet, accesses the PivotChart, disables its drop zones via PivotOptions, hides the chart legend, and saves the modified file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | PivotChart | hide legend | DropZonesVisible | PivotOptions | remove chart legend | Excel chart manipulation | save workbook | chart customization
// Common Searches: Aspose.Cells hide PivotChart legend C# | PivotOptions DropZonesVisible false Aspose.Cells | remove chart legend programmatically Aspose.Cells | disable pivot chart controls Aspose.Cells .NET | save workbook after chart changes Aspose.Cells
// Developer Intent: Load an existing Excel file, hide the legend and optional drop zones of a PivotChart, and save the updated workbook.
// Use Cases: Generate clean Excel reports where chart legends are unnecessary. | Create dashboards that prevent users from interacting with pivot chart controls. | Prepare workbooks for printing or PDF export by removing visual clutter. | Automate batch processing of multiple workbooks to standardize chart appearance.
// AI Prompts: Show C# code using Aspose.Cells to hide a PivotChart legend, disable drop zones, and save the workbook. | Provide an Aspose.Cells example that iterates over all worksheets and hides legends for every PivotChart. | Write a script that loads an Excel file, removes chart legends and pivot controls, and outputs a new file using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;

// Loads an Excel workbook, verifies a chart on the first worksheet, accesses the PivotChart, disables its drop zones via PivotOptions, hides the chart legend, and saves the modified file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one chart on the worksheet
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart on the sheet (assumed to be a PivotChart)
            Chart chart = worksheet.Charts[0];

            // Optionally hide all pivot controls on the chart
            chart.PivotOptions.DropZonesVisible = false;

            // Hide the chart legend using the Chart.ShowLegend property
            chart.ShowLegend = false;

            // Save the workbook with the changes
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
