// Title: Convert an Excel table to a range while preserving only the header and first data row using TableToRangeOptions.LastRow in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# convert ListObject to range keep header and first row only | TableToRangeOptions LastRow example preserving header formatting | How to retain only the first two rows of an Excel table when converting to range | C# code to convert Excel table to normal range while preserving header style | Aspose.Cells convert table to range with limited rows .NET
// Tags: TableToRangeOptions LastRow C# example | Aspose.Cells table to range conversion | preserve header style Excel Aspose.Cells | retain first row data Aspose.Cells | Excel table conversion .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;   // Required for ListObject and TableToRangeOptions

namespace AsposeCellsExample
{
    // The sample loads a workbook, selects the first ListObject, sets TableToRangeOptions.LastRow to include only the header and the first data row, converts the table to a normal range with those options, and saves the result to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one table (ListObject)
                if (worksheet.ListObjects.Count == 0)
                {
                    Console.WriteLine("No tables (ListObjects) found in the worksheet.");
                    return;
                }

                // Retrieve the first table on the worksheet
                ListObject table = worksheet.ListObjects[0];

                // Set conversion options to retain header row and first data row only
                TableToRangeOptions options = new TableToRangeOptions
                {
                    // Header row + first data row
                    LastRow = table.StartRow + 1
                };

                // Convert the table to a normal range using the specified options
                table.ConvertToRange(options);

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
