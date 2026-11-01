// Title: How to retrieve a chart from an Excel worksheet and list all its series names using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# get chart object from worksheet and read series names | list series names of the first chart in an Excel file using Aspose.Cells | enumerate NSeries collection in Aspose.Cells chart C# example | C# code to extract chart series labels from an .xlsx workbook with Aspose.Cells | how to iterate over chart series in Aspose.Cells for .NET
// Tags: Aspose.Cells retrieve chart from worksheet | Aspose.Cells enumerate chart series names | C# Aspose.Cells NSeries iteration | Excel .xlsx chart series extraction C# | Aspose.Cells chart series enumeration example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads an Excel workbook, accesses the first worksheet, checks for charts, retrieves the first chart, and iterates through its NSeries collection to output each series' Name, with error handling for missing files or absent charts.
class ListChartSeriesNames
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string workbookPath = "input.xlsx";

            // Verify that the file exists before loading
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"File not found: {workbookPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Retrieve the first chart
            Chart chart = worksheet.Charts[0];

            // Output the names of all series in the chart
            foreach (var series in chart.NSeries)
            {
                Console.WriteLine(series.Name);
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
