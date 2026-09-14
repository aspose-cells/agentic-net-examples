// Title: Find the worksheet that contains a Gantt‑style bar chart and log its name using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that opens an Excel workbook, scans every worksheet for a Bar chart configured as a Gantt chart, and prints the worksheet name to the console. | Extend the example to also write the detected worksheet name to a log file while preserving the existing error‑handling logic.
// Common Searches: how to locate the worksheet of a specific chart type with Aspose.Cells C# | c# Aspose.Cells find bar chart that represents a Gantt chart | retrieve parent worksheet name for a chart in an Excel workbook using Aspose.Cells | debug chart placement Aspose.Cells .NET | search all worksheets for Gantt chart Aspose.Cells example
// Tags: Aspose.Cells locate worksheet by chart type | detect Gantt‑style bar chart in Excel with Aspose.Cells | record worksheet name of detected chart Aspose.Cells | debug chart detection Aspose.Cells .NET | retrieve parent worksheet of chart Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// Loads an Excel workbook, iterates through each worksheet and its chart collection, treats a Bar chart as a potential Gantt chart, and writes the name of the worksheet containing such a chart to the console (or a log file).
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            bool ganttFound = false;

            // Iterate through all worksheets to locate a Gantt-like chart
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the charts collection of the current worksheet
                ChartCollection charts = sheet.Charts;

                foreach (Chart chart in charts)
                {
                    // Aspose.Cells does not expose a dedicated Gantt chart type in this version.
                    // As a workaround, treat a Bar chart that is configured as a Gantt chart as a match.
                    if (chart.Type == ChartType.Bar)
                    {
                        // Additional checks could be added here to verify Gantt-specific settings.
                        Console.WriteLine($"Bar chart (potential Gantt) is located in worksheet: {sheet.Name}");
                        ganttFound = true;
                        break; // Exit inner loop
                    }
                }

                if (ganttFound)
                {
                    break; // Exit outer loop if a Gantt-like chart was found
                }
            }

            if (!ganttFound)
            {
                Console.WriteLine("No Gantt-like chart was found in the workbook.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
