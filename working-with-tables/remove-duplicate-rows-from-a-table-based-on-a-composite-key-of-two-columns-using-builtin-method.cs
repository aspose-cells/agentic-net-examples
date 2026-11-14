// Title: Remove duplicate rows from an Excel worksheet using a composite key of two columns with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# remove duplicate rows based on two columns | How to use RemoveDuplicates with a composite key in Aspose.Cells | C# code to delete Excel rows that have the same values in column A and B | Eliminate duplicate records in an XLSX file using Aspose.Cells RemoveDuplicates method | Remove duplicate rows from a table after header row using Aspose.Cells .NET
// Tags: Aspose.Cells RemoveDuplicates method usage | duplicate row removal based on composite key C# | Excel worksheet duplicate elimination Aspose.Cells | remove duplicate rows Excel .NET API | composite key duplicate detection Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an XLSX workbook, determines the data range after the header row, and calls worksheet.Cells.RemoveDuplicates with parameters that define a two‑column composite key. Rows sharing identical values in those columns are removed, and the cleaned workbook is saved to a new file. Error handling covers missing input files and empty data ranges.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "Input.xlsx";
            string outputPath = "Output.xlsx";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                var workbook = new Workbook(inputPath);

                // Get the first worksheet
                var worksheet = workbook.Worksheets[0];

                // Settings for the composite key
                int headerRowIndex = 0;          // Header row (zero‑based)
                int keyFirstColumnIndex = 0;     // First column of the key (e.g., column A)
                int keyColumnCount = 2;          // Number of columns forming the key (e.g., A and B)

                // Determine data range
                int dataStartRow = headerRowIndex + 1; // First data row (zero‑based)
                int totalDataRows = worksheet.Cells.MaxDataRow - dataStartRow + 1;

                if (totalDataRows <= 0)
                {
                    Console.WriteLine("No data rows found to process.");
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to {outputPath}");
                    return;
                }

                // Remove duplicate rows based on the composite key columns
                // columnOffsets = null indicates contiguous columns starting at keyFirstColumnIndex
                worksheet.Cells.RemoveDuplicates(
                    dataStartRow,          // start row
                    keyFirstColumnIndex,  // start column
                    totalDataRows,        // total rows to process
                    keyColumnCount,       // total columns that constitute the key
                    false,                // case‑sensitive comparison flag
                    null);                // columnOffsets (null for default)

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
