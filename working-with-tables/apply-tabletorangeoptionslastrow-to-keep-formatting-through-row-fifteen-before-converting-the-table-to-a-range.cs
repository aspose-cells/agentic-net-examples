// Title: Convert an Excel table to a normal range while preserving formatting through row 15 using Aspose.Cells TableToRangeOptions in C#
// AI Prompts: Write C# code that loads a workbook, checks for a ListObject on the first worksheet, and uses TableToRangeOptions with LastRow = 15 to convert the table to a range while keeping its formatting. | Show how to apply Aspose.Cells TableToRangeOptions.LastRow to retain table styles up to a specific row when converting a table to a range in C#. | Create a snippet that validates the input file, converts the first Excel table to a range with formatting preserved through row 15, and saves the result.
// Common Searches: Aspose.Cells C# convert table to range keep formatting through row 15 | How to use TableToRangeOptions.LastRow to preserve Excel table styles | Convert ListObject to range with formatting limit Aspose.Cells example | C# code for converting Excel table to range while retaining first 15 rows formatting | Aspose.Cells preserve table formatting when converting to normal range
// Tags: Aspose.Cells TableToRangeOptions LastRow | convert Excel table to range C# | preserve table formatting Aspose.Cells | ListObject to range conversion | formatting retention up to specific row

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace Example
{
    // Loads an Excel workbook, verifies a table exists on the first worksheet, converts that table to a normal range using TableToRangeOptions with LastRow set to 15 to keep formatting through row 15, and saves the modified file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Ensure the workbook has at least one worksheet
                if (workbook.Worksheets.Count == 0)
                {
                    Console.WriteLine("The workbook contains no worksheets.");
                    return;
                }

                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one table (ListObject)
                if (worksheet.ListObjects.Count == 0)
                {
                    Console.WriteLine("No tables found on the first worksheet.");
                    return;
                }

                ListObject table = worksheet.ListObjects[0];

                // Configure conversion options to preserve formatting up to row 15
                TableToRangeOptions options = new TableToRangeOptions
                {
                    LastRow = 15
                };

                // Convert the table to a normal range using the specified options
                table.ConvertToRange(options);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
