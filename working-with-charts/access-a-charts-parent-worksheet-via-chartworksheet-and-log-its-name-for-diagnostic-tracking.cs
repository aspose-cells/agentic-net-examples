// Title: How to obtain and log a chart's parent worksheet name using the Chart.Worksheet property in Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get worksheet name from a chart object | using Chart.Worksheet to identify a chart's sheet in Aspose.Cells | log parent worksheet of a chart with Aspose.Cells .NET example | retrieve chart's parent sheet name via Aspose.Cells API | diagnostic logging of chart worksheet name in C# Aspose.Cells
// Tags: Chart.Worksheet property Aspose.Cells | chart parent worksheet retrieval C# | log chart worksheet name Aspose.Cells | diagnostic logging Excel chart sheet .NET | access chart worksheet via Aspose.Cells API

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an Excel workbook, accesses the first chart on the first worksheet, uses the Chart.Worksheet property to get the chart's parent worksheet, and writes the worksheet's name to the console, including file existence checks and exception handling.
class ChartParentWorksheetDemo
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Assume the first worksheet contains at least one chart
            Worksheet firstSheet = workbook.Worksheets[0];
            if (firstSheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the first worksheet.");
                return;
            }

            // Get the first chart on the worksheet
            Chart chart = firstSheet.Charts[0];

            // Access the chart's parent worksheet via Chart.Worksheet
            Worksheet parentWorksheet = chart.Worksheet;

            // Log the name of the parent worksheet
            Console.WriteLine($"Chart's parent worksheet name: {parentWorksheet.Name}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
