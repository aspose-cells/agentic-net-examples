// Title: How to keep Excel cell formatting when refreshing a QueryTable using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# preserve cell formatting when a QueryTable is refreshed | how to set PreserveFormatting = true on a QueryTable using Aspose.Cells .NET | maintain Excel cell styles after updating data with a QueryTable in C# | example code to enable formatting preservation for QueryTable in Aspose.Cells | prevent loss of cell formatting on QueryTable refresh with Aspose.Cells
// Tags: Aspose.Cells QueryTable PreserveFormatting property | keep Excel cell styles after QueryTable refresh | C# enable formatting preservation for query tables | save workbook with preserved formatting Aspose.Cells | refresh data without losing styles Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, accesses the first worksheet's QueryTable, sets its PreserveFormatting flag to true so that cell styles are retained during data refresh, and saves the modified workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook that contains the QueryTable
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one QueryTable
                if (worksheet.QueryTables.Count == 0)
                {
                    Console.WriteLine("No QueryTable found in the worksheet.");
                    return;
                }

                // Access the first QueryTable
                QueryTable queryTable = worksheet.QueryTables[0];

                // Preserve formatting after a data refresh
                queryTable.PreserveFormatting = true;

                // Save the workbook with the updated setting
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
