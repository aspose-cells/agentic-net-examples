// Title: Find the worksheet that hosts a Gantt chart and log its name using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that scans every worksheet in a workbook, identifies a chart whose Type is Gantt, and writes the worksheet name to the console. | Update an existing Aspose.Cells routine to exit loops once a Gantt chart is detected and output its parent worksheet for debugging.
// Common Searches: Aspose.Cells C# how to locate the worksheet containing a Gantt chart | retrieve sheet name of a specific chart type using Aspose.Cells .NET | debugging Aspose.Cells workbook by printing the parent worksheet of a Gantt chart | C# iterate through worksheets and charts to find Gantt chart in Excel file
// Tags: Aspose.Cells find worksheet by chart type | C# iterate worksheets and charts Aspose.Cells | log Gantt chart parent worksheet Aspose.Cells | debug Excel workbook chart location Aspose.Cells | detect Gantt chart type fallback Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an Excel workbook, loops through each worksheet and its charts, checks for a chart whose Type string equals "Gantt", captures the containing worksheet's name, and writes it to the console, with safeguards for missing files and runtime exceptions.
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

            // Variable to hold the name of the worksheet containing the Gantt chart
            string? ganttWorksheetName = null;

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all charts in the current worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    // Detect Gantt chart by name (fallback when ChartType.Gantt is unavailable)
                    if (chart.Type.ToString().Equals("Gantt", StringComparison.OrdinalIgnoreCase))
                    {
                        ganttWorksheetName = sheet.Name;
                        Console.WriteLine($"Gantt chart is located in worksheet: {ganttWorksheetName}");
                        break; // Exit inner loop
                    }
                }

                if (ganttWorksheetName != null)
                    break; // Exit outer loop
            }

            // If no Gantt chart was found, log a message
            if (ganttWorksheetName == null)
            {
                Console.WriteLine("No Gantt chart found in the workbook.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
