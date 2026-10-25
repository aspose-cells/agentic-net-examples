// Title: How to verify that a chart’s parent worksheet name matches an expected sheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get worksheet name of a chart | compare chart worksheet name with expected sheet in Aspose.Cells | check if a chart belongs to a specific worksheet using Aspose.Cells .NET | retrieve parent worksheet of first chart in Excel workbook with Aspose.Cells | validate chart location against expected sheet name in C#
// Tags: Aspose.Cells chart worksheet retrieval | C# verify chart parent sheet | Aspose.Cells workbook chart location validation | chart.Worksheet.Name property Aspose.Cells | Excel chart parent worksheet check .NET

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example loads an Excel workbook with Aspose.Cells, finds the first chart across all worksheets, reads the chart's containing worksheet via the chart.Worksheet.Name property, and compares it to a predefined expected sheet name, outputting whether the names match while handling missing files, absent charts, and runtime errors.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";

        // Ensure the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Expected parent worksheet name for the chart
            string expectedWorksheetName = "Sheet1";

            // Retrieve the first chart found in the workbook
            Chart chart = null;
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws.Charts.Count > 0)
                {
                    chart = ws.Charts[0];
                    break;
                }
            }

            if (chart == null)
            {
                Console.WriteLine("No chart found in the workbook.");
                return;
            }

            // Get the name of the worksheet that contains the chart
            string actualWorksheetName = chart.Worksheet.Name;

            // Verify that the retrieved worksheet name matches the expected name
            if (string.Equals(actualWorksheetName, expectedWorksheetName, StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("Worksheet name matches the expected chart parent worksheet.");
            }
            else
            {
                Console.WriteLine($"Mismatch: Expected '{expectedWorksheetName}', but chart is on worksheet '{actualWorksheetName}'.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
