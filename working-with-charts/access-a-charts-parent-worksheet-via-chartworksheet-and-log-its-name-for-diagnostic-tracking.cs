// Title: Retrieve and log a chart’s parent worksheet name with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an Excel workbook using Aspose.Cells, iterates through its charts, and writes each chart's parent worksheet name to the console. | Demonstrate how to access the Chart.Worksheet property in Aspose.Cells and log the worksheet's Name property for diagnostic purposes. | Create a C# snippet that checks for charts on a specific worksheet, retrieves the first chart, and prints the containing sheet's name.
// Common Searches: Aspose.Cells C# get worksheet name from chart object | How to use Chart.Worksheet to identify chart location in Excel with Aspose.Cells | Log parent sheet of a chart in Aspose.Cells .NET example | Iterate all charts in a workbook and output their worksheet names using Aspose.Cells
// Tags: Aspose.Cells Chart.Worksheet name retrieval | C# log chart parent worksheet Aspose.Cells | enumerate charts workbook Aspose.Cells | diagnostic worksheet name from chart Aspose.Cells | Excel chart location debugging Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an Excel workbook with Aspose.Cells, accesses a chart via Chart.Worksheet, and writes the parent worksheet's Name to the console for diagnostics.
class ChartWorksheetLogger
{
    static void Main()
    {
        string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the first worksheet.");
                return;
            }

            // Retrieve the first chart on the worksheet
            Chart chart = worksheet.Charts[0];

            // Obtain the chart's parent worksheet via Chart.Worksheet
            Worksheet parentSheet = chart.Worksheet;

            // Log the worksheet name for diagnostic tracking
            Console.WriteLine($"Chart's parent worksheet name: {parentSheet.Name}");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display an error message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
