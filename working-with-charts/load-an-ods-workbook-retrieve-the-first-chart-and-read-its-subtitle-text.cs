// Title: Read the subtitle (title) text of the first chart in an ODS workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an ODS file with Aspose.Cells, checks the first worksheet for charts, and prints the chart's title text as a subtitle. | Create a .NET example that validates the existence of an ODS workbook, loads it, and safely extracts the subtitle of the first chart, handling cases where no chart is present. | Generate a robust C# snippet using Aspose.Cells that reads the subtitle (title) of the first chart in an OpenDocument spreadsheet and logs it, with proper error handling for missing files.
// Common Searches: Aspose.Cells C# get subtitle of first chart in ODS file | How to read chart title text from an OpenDocument spreadsheet using Aspose.Cells .NET | C# example for loading ODS workbook and extracting chart metadata with Aspose.Cells | Retrieve chart subtitle from ODS workbook in C# Aspose.Cells
// Tags: load ODS workbook Aspose.Cells .NET | retrieve first chart subtitle C# | chart title property Aspose.Cells | handle missing chart ODS spreadsheet | validate file existence Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample checks that 'input.ods' exists, loads it into an Aspose.Cells Workbook, accesses the first worksheet, verifies that at least one chart is present, retrieves the first chart, reads its Title.Text (used as a subtitle placeholder), outputs the text, and includes error handling for missing files or other runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputFile = "input.ods";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Input file '{inputFile}' was not found.");
            return;
        }

        try
        {
            // Load the ODS workbook
            Workbook workbook = new Workbook(inputFile);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count > 0)
            {
                // Retrieve the first chart
                Chart chart = sheet.Charts[0];

                // Read the chart title text (Aspose.Cells does not expose a separate subtitle property)
                string titleText = chart.Title?.Text ?? string.Empty;

                // Output the title (used as subtitle placeholder)
                Console.WriteLine("Chart title (used as subtitle): " + titleText);
            }
            else
            {
                Console.WriteLine("No charts found in the first worksheet.");
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
