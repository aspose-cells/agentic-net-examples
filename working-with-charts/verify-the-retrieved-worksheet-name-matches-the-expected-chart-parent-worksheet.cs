// Title: Check if a chart's parent worksheet name matches a specific sheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that opens an Excel workbook, retrieves the first chart, reads its Worksheet.Name property, and compares it to a supplied sheet name. | Create a C# snippet that prints whether the chart resides on a worksheet named "DataSheet" and then saves the workbook without modifications.
// Common Searches: Aspose.Cells C# how to find the worksheet that contains a chart | compare chart parent worksheet name with expected sheet name using Aspose.Cells | verify chart location in an Excel file with Aspose.Cells .NET example | retrieve chart's worksheet name and validate against a string in C#
// Tags: Aspose.Cells get chart parent worksheet name | Aspose.Cells compare worksheet name with expected value | C# validate chart location in Excel workbook | Aspose.Cells chart worksheet verification

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The program loads "input.xlsx" with Aspose.Cells, accesses the first worksheet and its first chart, obtains the chart's parent worksheet name, compares it to the expected name "DataSheet", outputs a match or mismatch message, and saves the workbook as "output.xlsx".
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the first worksheet.");
                return;
            }

            // Retrieve the first chart
            Chart chart = worksheet.Charts[0];

            // Get the name of the worksheet that is the parent of the chart
            string chartParentName = chart.Worksheet.Name;

            // Define the expected worksheet name
            string expectedWorksheetName = "DataSheet";

            // Verify that the retrieved worksheet name matches the expected name
            if (chartParentName.Equals(expectedWorksheetName, StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("Chart parent worksheet name matches the expected worksheet.");
            }
            else
            {
                Console.WriteLine($"Mismatch: Chart is in worksheet '{chartParentName}', expected '{expectedWorksheetName}'.");
            }

            // Save the workbook (no changes made, but follows the save rule)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
