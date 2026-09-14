// Title: Remove duplicate rows from a CSV file using Aspose.Cells in C# and export to XLSX
// AI Prompts: Generate C# code that loads a CSV with Aspose.Cells, removes rows that have duplicate values in a specified column, and saves the cleaned data as an XLSX workbook. | Adjust the duplicate‑removal step to keep the last occurrence of each key column value when calling Cells.RemoveDuplicates. | Extend the solution to accept a custom delimiter and perform case‑insensitive duplicate detection while processing the CSV with Aspose.Cells.
// Common Searches: aspnet remove duplicate rows from csv using aspose.cells | c# aspose.cells import csv and delete duplicate entries based on column | how to use Cells.RemoveDuplicates with key column index in Aspose.Cells | convert csv to xlsx while filtering duplicate records in .NET
// Tags: Aspose.Cells duplicate row removal | Cells.RemoveDuplicates key column usage | CSV to XLSX conversion with duplicate filtering | C# duplicate detection in CSV using Aspose.Cells | Aspose.Cells CSV data cleaning

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDuplicateRemoval
{
    // Loads a CSV into an Aspose.Cells workbook, removes rows that share duplicate values in a designated key column using Cells.RemoveDuplicates, and saves the deduplicated result as an XLSX file.
    public class RemoveCsvDuplicates
    {
        public static void Run()
        {
            try
            {
                // Path to the source CSV file
                string csvPath = "input.csv";

                // Path for the resulting Excel file
                string outputPath = "output.xlsx";

                // Index of the column that serves as the key for duplicate detection (0‑based)
                int keyColumnIndex = 0; // Example: first column

                // Verify that the CSV file exists
                if (!File.Exists(csvPath))
                {
                    Console.WriteLine($"CSV file not found: {csvPath}");
                    return;
                }

                // Create a new workbook and get the first worksheet's cells collection
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Import the CSV data starting at cell A1 (row 0, column 0)
                cells.ImportCSV(csvPath, ",", true, 0, 0);

                // Determine the actual range of imported data
                int lastRow = cells.MaxDataRow;          // zero‑based index of the last row with data
                int lastColumn = cells.MaxDataColumn;    // zero‑based index of the last column with data

                // Remove duplicate rows based on the key column
                cells.RemoveDuplicates(0, 0, lastRow, lastColumn, true, new int[] { keyColumnIndex });

                // Save the cleaned workbook as an XLSX file
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Duplicates removed. Output saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            RemoveCsvDuplicates.Run();
        }
    }
}
