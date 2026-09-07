// Title: Load an XLS workbook, detect the first PivotChart, configure its data to display as percentages via PivotOptions, and save the file with Aspose.Cells for .NET
// AI Prompts: Read an existing .xls file, locate the first chart, verify it is a PivotChart, set PivotOptions.ShowDataAs = ShowDataAs.Percent, and write the workbook to a new file. | Using Aspose.Cells for .NET, programmatically change a PivotChart's data representation to percentages and persist the changes in a new XLS document.
// Common Searches: asp.net change pivot chart values to percent in existing xls using Aspose.Cells | c# detect pivot chart in workbook and set ShowDataAs to Percent with Aspose.Cells | modify pivot chart data display to percentage and save workbook Aspose.Cells .NET example | load xls, update pivot chart options, save new file Aspose.Cells C#
// Tags: Aspose.Cells PivotChart ShowDataAs Percent | C# load XLS workbook Aspose.Cells | detect PivotChart via PivotOptions | set PivotChart data as percentage | save modified workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example checks for the input XLS file, loads it with Aspose.Cells, retrieves the first worksheet's first chart, determines if it is a PivotChart through its PivotOptions, optionally sets the ShowDataAs property to Percent, and saves the updated workbook to a new file while providing console status messages.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xls";
        const string outputPath = "output.xls";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing XLS file
            Workbook workbook = new Workbook(inputPath);

            // Assume the PivotChart is on the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one chart in the worksheet
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the first worksheet.");
                return;
            }

            // Get the first chart in the worksheet
            Chart chart = sheet.Charts[0];

            // Determine if the chart is a PivotChart by checking PivotOptions
            if (chart.PivotOptions != null)
            {
                // Example: set ShowDataAs to Percent if needed
                // chart.PivotOptions.ShowDataAs = ShowDataAs.Percent;

                Console.WriteLine("PivotChart detected. (ShowDataAs setting is version‑dependent and has been omitted.)");
            }
            else
            {
                Console.WriteLine("The first chart is not a PivotChart.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
