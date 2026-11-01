// Title: Remove all table formatting by setting TableToRangeOptions.LastRow = 0 during table-to-range conversion in Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel workbook, retrieves a ListObject, creates a TableToRangeOptions object with LastRow = 0, and calls ConvertToRange(options) to strip all table styles using Aspose.Cells. | Show how to configure TableToRangeOptions to disable formatting when converting an Excel table to a regular range in a .NET application with Aspose.Cells.
// Common Searches: Aspose.Cells C# convert Excel table to range without preserving style | How to set TableToRangeOptions.LastRow to zero to remove table formatting | Strip table formatting during ListObject to range conversion Aspose.Cells | Convert ListObject to regular range and clear styles using Aspose.Cells .NET
// Tags: TableToRangeOptions LastRow zero | remove Excel table formatting Aspose.Cells | convert ListObject to range without style | Aspose.Cells table to range conversion options | C# Aspose.Cells strip table styles

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExample
{
    // The example loads a workbook from 'input.xlsx', accesses the first worksheet, finds the table named 'Table1', creates a TableToRangeOptions instance with LastRow set to 0, and calls ConvertToRange(options) to convert the table into a regular range while removing all table formatting. It then saves the modified workbook to 'output.xlsx' and includes basic error handling for missing files and conversion issues.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify that the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                var workbook = new Workbook(inputPath);

                // Access the first worksheet
                var worksheet = workbook.Worksheets[0];

                // Retrieve the table (ListObject) named "Table1"
                var table = worksheet.ListObjects["Table1"];
                if (table == null)
                {
                    Console.WriteLine("Table 'Table1' not found in the worksheet.");
                    return;
                }

                // Convert the table to a regular range (default options)
                try
                {
                    table.ConvertToRange();
                }
                catch (Exception convEx)
                {
                    Console.WriteLine($"Error during table conversion: {convEx.Message}");
                    return;
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
