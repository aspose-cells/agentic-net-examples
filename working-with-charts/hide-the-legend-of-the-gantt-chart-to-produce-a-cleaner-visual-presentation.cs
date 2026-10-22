// Title: Hide Gantt chart legend in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Set the Chart.ShowLegend property to false for a Gantt chart and save the workbook with Aspose.Cells. | Load an existing .xlsx file, locate the first chart, disable its legend, and write the updated file using C#. | Programmatically remove the legend from a worksheet chart without affecting other chart elements in Aspose.Cells.
// Common Searches: C# Aspose.Cells hide legend on Gantt chart in existing Excel file | How to remove chart legend from first worksheet chart using Aspose.Cells .NET | Aspose.Cells example to disable legend for Excel Gantt chart programmatically
// Tags: Aspose.Cells chart ShowLegend property C# | hide Gantt chart legend Aspose.Cells | modify Excel chart legend .NET | remove chart legend from .xlsx using Aspose.Cells | programmatic chart formatting Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads an existing Excel workbook, accesses the first worksheet and its first chart (assumed to be a Gantt chart), sets ShowLegend to false to hide the legend, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook that contains the Gantt chart
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the Gantt chart (assumed to be the first chart in the sheet)
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("Error: No charts found in the worksheet.");
                return;
            }

            Chart ganttChart = worksheet.Charts[0];

            // Hide the legend to achieve a cleaner visual presentation
            ganttChart.ShowLegend = false;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
