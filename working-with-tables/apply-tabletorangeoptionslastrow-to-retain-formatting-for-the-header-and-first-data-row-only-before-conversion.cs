// Title: Use TableToRangeOptions.LastRow in Aspose.Cells for .NET to convert an Excel table to a range while keeping header and first data row formatting
// Common Searches: Aspose.Cells TableToRangeOptions.LastRow keep header formatting during table conversion | C# convert Excel ListObject to range preserving only first data row style | how to retain header and first row formatting when converting a table to a range with Aspose.Cells | selective formatting options for table‑to‑range conversion in Aspose.Cells .NET | example of using TableToRangeOptions to limit formatting to header and first row in C#
// Tags: TableToRangeOptions.LastRow usage Aspose.Cells | convert ListObject to range preserving header style | retain first data row formatting C# Aspose.Cells | selective formatting during Excel table conversion | Aspose.Cells .NET table to range conversion options

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, accesses the first worksheet's ListObject, and converts the table to a normal range using TableToRangeOptions.LastRow so that only the header row and the first data row keep their original formatting, then saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException($"Input file not found: {inputPath}");

                // Load the existing workbook
                var workbook = new Workbook(inputPath);

                // Access the first worksheet (or any worksheet containing the table)
                var worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one table (ListObject)
                if (worksheet.ListObjects.Count == 0)
                    throw new InvalidOperationException("No tables (ListObjects) found in the worksheet.");

                // Assume the table to be converted is the first ListObject in the worksheet
                var table = worksheet.ListObjects[0];

                // Convert the table to a normal range (preserves data; formatting is retained by default)
                table.ConvertToRange();

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log the exception details for troubleshooting
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
