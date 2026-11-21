// Title: Load an ODS workbook and read the first chart's subtitle (title) text with Aspose.Cells for .NET
// Common Searches: how to get the subtitle of the first chart in an ODS file using Aspose.Cells C# | Aspose.Cells C# read chart title text from ODS workbook | retrieve chart subtitle from ODS spreadsheet with Aspose.Cells .NET | C# example to load ODS workbook and access chart title using Aspose.Cells | read chart title from ODS file programmatically with Aspose.Cells for .NET
// Tags: load ODS workbook with Aspose.Cells | retrieve first chart from worksheet Aspose.Cells | extract chart title text C# | read chart subtitle from ODS file Aspose.Cells | Aspose.Cells ODS chart processing .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample verifies that an input.ods file exists, loads it into an Aspose.Cells Workbook, accesses the first worksheet, checks for charts, retrieves the first Chart object, reads its Title.Text (used as the subtitle), and prints the subtitle (or title) text to the console while handling potential errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.ods";

        try
        {
            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"File not found: {inputPath}");
                return;
            }

            // Load the ODS workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (charts are stored in worksheets)
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the first chart in the worksheet
            if (sheet.Charts.Count > 0)
            {
                Chart chart = sheet.Charts[0];

                // Read the chart title text (Aspose.Cells does not expose a separate subtitle property in all versions)
                string subtitleText = chart.Title != null ? chart.Title.Text : string.Empty;

                // Output the subtitle (or title) text
                Console.WriteLine("Chart Subtitle (or Title): " + subtitleText);
            }
            else
            {
                Console.WriteLine("No charts found in the first worksheet.");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
