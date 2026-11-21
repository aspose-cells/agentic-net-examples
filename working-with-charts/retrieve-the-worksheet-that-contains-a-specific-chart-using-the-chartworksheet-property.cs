// Title: Get the worksheet that contains a specific chart using the Chart.Worksheet property in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# get worksheet of a chart | Chart.Worksheet property example in .NET | Find parent worksheet for a chart using Aspose.Cells | Determine which sheet a chart belongs to with Aspose.Cells C# | Retrieve chart owner worksheet from a workbook in C#
// Tags: Chart.Worksheet property Aspose.Cells | retrieve chart parent worksheet C# | determine chart sheet Aspose.Cells | access chart's worksheet .NET | Aspose.Cells chart ownership

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an Excel workbook, checks for at least one worksheet and chart, obtains the first chart, uses the Chart.Worksheet property to identify the worksheet that owns the chart, prints the worksheet name, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure there is at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("The workbook contains no worksheets.");
                return;
            }

            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("The worksheet contains no charts.");
                return;
            }

            // Retrieve the first chart
            Chart chart = worksheet.Charts[0];

            // Get the worksheet that owns this chart
            Worksheet chartOwnerWorksheet = chart.Worksheet;

            // Display the worksheet name
            Console.WriteLine("The chart is located on worksheet: " + chartOwnerWorksheet.Name);

            // Save the workbook (optional)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
