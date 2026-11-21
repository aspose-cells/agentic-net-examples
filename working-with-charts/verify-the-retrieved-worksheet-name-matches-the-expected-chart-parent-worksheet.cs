// Title: How to verify a chart's parent worksheet name matches an expected sheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# get worksheet name that contains a specific chart | C# code to compare chart parent sheet name with expected value using Aspose.Cells | Check if a chart is placed on Sheet1 with Aspose.Cells .NET | Retrieve the worksheet of the first chart in an Excel workbook with Aspose.Cells | Validate chart location against expected worksheet in C# Aspose.Cells example | How to use chart.Worksheet.Name property in Aspose.Cells for .NET
// Tags: chart.Worksheet.Name Aspose.Cells | chart parent sheet identification Aspose.Cells | retrieve first chart worksheet Aspose.Cells | compare expected worksheet name Aspose.Cells | locate chart in workbook Aspose.Cells .NET | chart sheet validation example Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample loads an Excel workbook with Aspose.Cells, finds the first chart, reads its containing worksheet via chart.Worksheet.Name, and compares that name to a predefined expected sheet, outputting whether the chart resides on the correct worksheet.
class ChartParentVerification
{
    static void Main()
    {
        string inputPath = "input.xlsx";

        // Verify that the input file exists before loading
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Expected parent worksheet name for the chart
            string expectedWorksheetName = "Sheet1";

            // Locate the first chart in the workbook
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

            // Retrieve the name of the worksheet that contains the chart
            string parentWorksheetName = chart.Worksheet.Name;

            // Verify that the retrieved worksheet name matches the expected name
            if (parentWorksheetName.Equals(expectedWorksheetName, StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("Chart parent worksheet name matches expected: " + parentWorksheetName);
            }
            else
            {
                Console.WriteLine($"Mismatch: Chart is on worksheet '{parentWorksheetName}', expected '{expectedWorksheetName}'.");
            }

            // Save the workbook if further processing is required
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
