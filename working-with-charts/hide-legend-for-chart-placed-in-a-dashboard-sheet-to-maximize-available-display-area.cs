// Title: Hide the legend of the first chart on a Dashboard worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that opens an existing workbook, selects the 'Dashboard' sheet, and sets ShowLegend = false on its first chart before saving. | Show how to check for charts on a specific worksheet and programmatically hide their legends using the Aspose.Cells Chart.ShowLegend property in C#. | Provide a complete example that validates the presence of a chart on a dashboard sheet, disables the legend, and saves the updated Excel file with Aspose.Cells.
// Common Searches: aspnet hide legend first chart dashboard sheet Aspose.Cells | C# Aspose.Cells remove chart legend from Excel dashboard | how to disable chart legend on a specific worksheet using Aspose.Cells | Aspose.Cells chart.ShowLegend false example | programmatically hide Excel chart legend in .NET
// Tags: Aspose.Cells chart ShowLegend property C# | hide Excel chart legend Aspose.Cells | dashboard worksheet chart customization Aspose.Cells | C# remove chart legend from workbook | Aspose.Cells first chart legend suppression

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads 'Dashboard.xlsx', accesses the worksheet named 'Dashboard', checks for existing charts, and sets ShowLegend to false on the first chart to hide its legend, then saves the workbook as 'Dashboard_Updated.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "Dashboard.xlsx";
            const string outputPath = "Dashboard_Updated.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the dashboard worksheet
            Worksheet dashboardSheet = workbook.Worksheets["Dashboard"];
            if (dashboardSheet == null)
            {
                Console.WriteLine("Worksheet 'Dashboard' not found.");
                return;
            }

            // Hide legend of the first chart, if any
            if (dashboardSheet.Charts.Count > 0)
            {
                Chart chart = dashboardSheet.Charts[0];
                chart.ShowLegend = false;
            }
            else
            {
                Console.WriteLine("No charts found on the Dashboard sheet.");
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
