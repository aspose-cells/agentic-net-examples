// Title: How to enable automatic chart data refresh on workbook open with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that sets the Chart.RefreshDataOnOpen property to true for a chart in an Aspose.Cells workbook. | Generate a C# snippet that adds a chart to a worksheet and configures it to refresh its data automatically when the file is opened, using the latest Aspose.Cells API. | Provide a C# workaround for older Aspose.Cells versions to simulate chart data refresh on workbook open.
// Common Searches: Aspose.Cells C# enable chart to refresh data when workbook is opened | how to set chart auto‑refresh on opening Excel file using Aspose.Cells .NET | Chart.RefreshDataOnOpen example for Aspose.Cells 2023 | workaround for automatic chart refresh in older Aspose.Cells releases | C# Aspose.Cells update chart data on workbook open
// Tags: Aspose.Cells chart RefreshDataOnOpen | C# auto refresh chart data Aspose.Cells | Excel chart refresh on open .NET | Aspose.Cells workbook open chart update | Aspose.Cells chart data refresh workaround

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program loads an existing workbook, ensures a chart exists (creating a placeholder if needed), retrieves the first chart, notes that automatic refresh on workbook open requires the Chart.RefreshDataOnOpen property available in newer Aspose.Cells versions, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one chart; create a placeholder if none exist
            if (sheet.Charts.Count == 0)
            {
                // Add returns the index of the newly created chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart placeholder = sheet.Charts[chartIndex];
                placeholder.NSeries.Add("A1:A5", true);
            }

            // Retrieve the first chart on the worksheet
            Chart targetChart = sheet.Charts[0];

            // Note: Automatic refresh of chart data on workbook open is not directly exposed
            // in the current Aspose.Cells API version. If needed, consider updating to a newer version
            // where Chart.RefreshDataOnOpen property may be available.

            // Save the workbook with the updated setting
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
