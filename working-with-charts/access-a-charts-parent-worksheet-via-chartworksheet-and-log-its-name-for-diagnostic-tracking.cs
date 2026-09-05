// Title: How to get and log a chart's parent worksheet name with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an Excel workbook, retrieves the first chart, accesses its Worksheet property, and prints the worksheet name to the console. | Show an example of using Aspose.Cells Chart.Worksheet to identify the sheet containing a specific chart and log that name for diagnostics. | Provide a snippet that checks for charts on a worksheet, obtains the chart's parent worksheet, and writes the sheet name to the output.
// Common Searches: Aspose.Cells C# get worksheet name from Chart.Worksheet property | How to log the sheet that contains a chart using Aspose.Cells .NET | Retrieve parent worksheet of a chart in an existing Excel workbook with Aspose.Cells | C# code to print chart's parent worksheet name using Aspose.Cells | Aspose.Cells Chart.Worksheet example for diagnostic logging
// Tags: Aspose.Cells Chart.Worksheet property | retrieve parent worksheet of chart | log chart container sheet name C# | access chart's worksheet .NET | diagnostic logging of chart parent sheet Aspose.Cells | Excel chart parent worksheet extraction using Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example loads an existing Excel file, verifies that a chart exists, accesses the first chart's parent worksheet via the Chart.Worksheet property, prints the worksheet's Name to the console for diagnostic purposes, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Retrieve the first chart
            Chart chart = worksheet.Charts[0];

            // Access the chart's parent worksheet via Chart.Worksheet
            Worksheet parentWorksheet = chart.Worksheet;

            // Log the name of the parent worksheet
            Console.WriteLine($"Chart's parent worksheet name: {parentWorksheet.Name}");

            // Save the workbook (if any changes were made)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
