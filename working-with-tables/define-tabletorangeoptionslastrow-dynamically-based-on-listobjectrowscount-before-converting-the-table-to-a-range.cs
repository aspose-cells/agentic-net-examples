// Title: Set TableToRangeOptions.LastRow dynamically from a ListObject’s row count before converting an Excel table to a range with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# calculate TableToRangeOptions.LastRow using ListObject.Rows.Count | convert Excel table to range with dynamic last row in Aspose.Cells .NET | determine number of data rows in a ListObject before calling ConvertToRange in C# | C# example for setting TableToRangeOptions.LastRow based on table row count | programmatically convert Excel ListObject to range with variable last row Aspose.Cells
// Tags: TableToRangeOptions LastRow dynamic calculation | Aspose.Cells ListObject row count | convert Excel table to range C# Aspose.Cells | dynamic table conversion options Aspose.Cells | C# Excel ListObject data rows count

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExample
{
    // The example loads a workbook, verifies a ListObject exists, computes the number of data rows (excluding the header) from the first table, assigns that count to TableToRangeOptions.LastRow, converts the table to a range using those options, and saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                var workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index as needed)
                var worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one table
                if (worksheet.ListObjects.Count == 0)
                {
                    Console.WriteLine("No tables (ListObjects) found in the worksheet.");
                    return;
                }

                // Retrieve the first table (ListObject) on the worksheet
                var listObject = worksheet.ListObjects[0];

                // Determine the number of data rows in the table (excluding the header row)
                int dataRowCount = listObject.DataRange.RowCount - 1;
                if (dataRowCount < 0) dataRowCount = 0;

                // Create TableToRangeOptions and set LastRow dynamically based on the table's row count
                var options = new TableToRangeOptions
                {
                    // LastRow is zero‑based relative to the table's start row
                    LastRow = dataRowCount
                };

                // Convert the table to a range using the configured options
                listObject.ConvertToRange(options);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
