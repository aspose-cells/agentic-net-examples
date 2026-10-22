// Title: How to bind data label number format to the first chart series using Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# program that loads an Excel workbook with Aspose.Cells, retrieves the first chart, and sets series[0].DataLabels.NumberFormatLinked = true. | Generate code to enable linked number formatting for data labels of the first series in a chart using Aspose.Cells for .NET. | Create a snippet that opens an existing .xlsx file, finds the first chart series, turns on DataLabels.NumberFormatLinked, and saves the changes.
// Common Searches: Aspose.Cells set DataLabels.NumberFormatLinked for chart series in C# | How to bind data label formatting to series number format in Excel using Aspose.Cells | C# example linking chart series number format to data labels with Aspose.Cells | Enable linked number format for data labels of the initial chart series in .NET
// Tags: Aspose.Cells chart series DataLabels.NumberFormatLinked | bind data label number format to series Aspose.Cells C# | first chart series data label formatting .NET | link chart data labels to series number format Aspose | modify Excel chart data labels Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program loads an existing Excel workbook, accesses the first worksheet and its first chart, sets the DataLabels.NumberFormatLinked property of the chart's first series to true (binding data label formatting to the series number format), and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            var workbook = new Workbook(inputPath);

            // Get the first worksheet
            var worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart on the worksheet
            var chart = worksheet.Charts[0];

            // Ensure the chart has at least one series
            if (chart.NSeries.Count == 0)
            {
                Console.WriteLine("The chart does not contain any series.");
                return;
            }

            // Get the first series of the chart
            var series = chart.NSeries[0];

            // Bind the data label formatting to the series number format
            series.DataLabels.NumberFormatLinked = true;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
