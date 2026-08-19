// Title: Retrieve the subtitle of the first chart in an ODS workbook using Aspose.Cells for .NET (C#)
// Description: A concise C# example that checks for an ODS file, loads it with Aspose.Cells, accesses the first worksheet, verifies chart presence, reads the SubTitle property of the first chart (with null safety), and writes the subtitle text to the console.
// Keywords: Aspose.Cells | ODS workbook | chart subtitle | C# | .NET | read chart SubTitle | load ODS file | first chart access | worksheet charts
// Common Searches: Aspose.Cells read chart subtitle ODS C# | how to get chart subtitle from ODS using .NET | C# example load ODS workbook and retrieve first chart subtitle | Aspose.Cells SubTitle property ODS format
// Developer Intent: Extract the subtitle text of the first chart in an ODS file.
// Use Cases: Validate chart subtitles against corporate naming standards before publishing the workbook. | Display chart subtitles in a custom reporting dashboard after loading an ODS file. | Log all chart subtitles from an ODS workbook for compliance auditing.
// AI Prompts: Generate C# code with Aspose.Cells that iterates through every chart in an ODS workbook and prints each chart's title and subtitle, handling missing subtitles gracefully. | Provide a robust example that loads an ODS file, checks for chart existence, and returns the subtitle of a chart at a given index, including error handling for missing files and empty chart collections.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// A concise C# example that checks for an ODS file, loads it with Aspose.Cells, accesses the first worksheet, verifies chart presence, reads the SubTitle property of the first chart (with null safety), and writes the subtitle text to the console.
class Program
{
    static void Main()
    {
        const string inputPath = "input.ods";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the ODS workbook from file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts were found in the first worksheet.");
                return;
            }

            // Retrieve the first chart in the worksheet's chart collection
            Chart chart = worksheet.Charts[0];

            // Read the subtitle text of the chart (ODS format supports SubTitle)
            string subtitleText = chart.SubTitle?.Text ?? string.Empty;

            // Output the subtitle text
            Console.WriteLine("Chart Subtitle: " + subtitleText);
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
