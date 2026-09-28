// Title: Use TableToRangeOptions.LastRow to preserve formatting through row 15 when converting an Excel table to a range with Aspose.Cells in C#
// AI Prompts: Load a workbook, set TableToRangeOptions.LastRow = 14, and call ListObject.ConvertToRange to keep formatting up to row 15. | Write C# code that checks for a ListObject, applies TableToRangeOptions to retain cell styles for the first fifteen rows, and saves the workbook. | Show how to limit the formatting retention range during a table‑to‑range conversion using Aspose.Cells TableToRangeOptions.
// Common Searches: Aspose.Cells C# TableToRangeOptions LastRow keep formatting through row 15 | how to convert Excel table to range while preserving first 15 rows styles using Aspose | C# example limiting table-to-range conversion rows with Aspose.Cells | retain table formatting up to a specific row when converting to range in Aspose.Cells | convert ListObject to normal range and keep formatting for rows 1-15 Aspose.Cells
// Tags: Aspose.Cells TableToRangeOptions row limit | keep table styles after conversion to range | convert ListObject to normal range C# | limit formatting rows during table-to-range conversion | retain cell styles up to row fifteen Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExample
{
    // The example loads input.xlsx, verifies a table exists in the first worksheet, sets TableToRangeOptions.LastRow to 14 to keep formatting through row 15, converts the first ListObject to a normal range, and saves the modified workbook as output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one table (ListObject)
                if (sheet.ListObjects.Count == 0)
                {
                    Console.WriteLine("No tables (ListObjects) found in the worksheet.");
                    return;
                }

                // Get the first table
                ListObject table = sheet.ListObjects[0];

                // Set conversion options (keep formatting up to row 15, zero‑based index 14)
                TableToRangeOptions options = new TableToRangeOptions
                {
                    LastRow = 14
                };

                // Convert the table to a normal range
                table.ConvertToRange(options);

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
