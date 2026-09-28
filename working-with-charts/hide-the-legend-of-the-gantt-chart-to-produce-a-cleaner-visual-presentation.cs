// Title: Hide the legend of a Gantt chart in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Load an Excel workbook, locate the Gantt chart, set ShowLegend = false, and save the file with Aspose.Cells in C#. | Programmatically remove the legend from the first chart on a worksheet using the Aspose.Cells Chart.ShowLegend property. | Update an existing Excel file to hide chart legends without recreating the chart, leveraging the Aspose.Cells .NET API.
// Common Searches: Aspose.Cells C# hide legend on Gantt chart Excel file | remove chart legend from existing workbook using Aspose.Cells .NET | how to disable legend for first chart in worksheet with Aspose.Cells | C# Aspose.Cells set ShowLegend false for Excel chart | example code to hide Excel chart legend using Aspose.Cells library
// Tags: Aspose.Cells chart legend removal | C# ShowLegend property usage | Excel Gantt chart styling Aspose.Cells | modify chart legend visibility Aspose.Cells | Aspose.Cells chart API hide legend

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The code loads GanttChart.xlsx, accesses the first chart on the first worksheet, disables its legend by setting ShowLegend to false, and saves the modified workbook as GanttChart_NoLegend.xlsx.
class Program
{
    static void Main()
    {
        const string inputPath = "GanttChart.xlsx";
        const string outputPath = "GanttChart_NoLegend.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that contains the Gantt chart
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (modify the index if the chart is on a different sheet)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one chart on the worksheet
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the worksheet.");
                return;
            }

            // Retrieve the first chart on the worksheet (assumed to be the Gantt chart)
            Chart ganttChart = worksheet.Charts[0];

            // Hide the legend for a cleaner visual presentation
            ganttChart.ShowLegend = false;

            // Save the workbook with the updated chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
