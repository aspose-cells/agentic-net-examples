// Title: Retrieve the worksheet that hosts a Gantt chart and log its name using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells find which worksheet contains a specific chart | How to output the sheet name that holds a Gantt chart in Aspose.Cells | Example code to access Chart.Worksheet property in Aspose.Cells .NET | Debugging Excel Gantt chart location with Aspose.Cells C# | Retrieve hosting worksheet for the first chart in a workbook using Aspose.Cells
// Tags: Aspose.Cells chart worksheet property | C# retrieve chart host sheet | log worksheet name Aspose.Cells | debug Gantt chart location .NET | access Chart.Worksheet Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program loads an Excel workbook, locates the first chart (assumed to be a Gantt chart), obtains the Worksheet object that contains the chart via the Chart.Worksheet property, and writes the worksheet's Name to the console, with error handling for missing files or charts.
class Program
{
    static void Main(string[] args)
    {
        // Determine the input file path (argument or default)
        string inputPath = args.Length > 0 ? args[0] : "input.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Locate the first chart (assumed to be the Gantt chart)
            Chart ganttChart = null;
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    ganttChart = chart; // take the first chart found
                    break;
                }
                if (ganttChart != null)
                    break;
            }

            // Handle case where no chart is present
            if (ganttChart == null)
            {
                Console.WriteLine("No chart was found in the workbook.");
                return;
            }

            // Retrieve the worksheet that hosts the chart
            Worksheet hostingWorksheet = ganttChart.Worksheet;

            // Output the worksheet name for debugging
            Console.WriteLine($"Gantt chart is hosted on worksheet: {hostingWorksheet.Name}");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
