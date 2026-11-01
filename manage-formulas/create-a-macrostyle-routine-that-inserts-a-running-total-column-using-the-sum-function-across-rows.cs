// Title: Create a macro‑style C# routine to add a running total column with cumulative SUM formulas using Aspose.Cells
// Common Searches: C# code to add a running total column in an Excel file with Aspose.Cells | How to programmatically insert cumulative SUM formulas in a worksheet using Aspose.Cells for .NET | Aspose.Cells example for generating a running total column next to source data | Insert a running total column with header support using Aspose.Cells C# API | Automate Excel running total calculation across rows with Aspose.Cells library
// Tags: Aspose.Cells cumulative sum column | C# insert running total column | Excel worksheet dynamic SUM formula | header aware total column Aspose.Cells | macro style routine Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example defines a static C# method that loads an Excel workbook with Aspose.Cells, determines the last used row, optionally adds a "Running Total" header, computes the source column letter, and fills the adjacent column with SUM formulas that cumulatively total the source values from the first data row to each row, then saves the modified workbook.
    public class RunningTotalInserter
    {
        /// <param name="inputFilePath">Path to the source Excel file.</param>
        /// <param name="outputFilePath">Path where the modified workbook will be saved.</param>
        /// <param name="sourceColumnIndex">Zero‑based index of the column that contains the values to total (e.g., 0 for column A).</param>
        /// <param name="headerRowIndex">Zero‑based index of the header row. Set to -1 if there is no header.</param>
        public static void InsertRunningTotal(string inputFilePath, string outputFilePath, int sourceColumnIndex = 0, int headerRowIndex = 0)
        {
            // Verify input file exists.
            if (!File.Exists(inputFilePath))
                throw new FileNotFoundException($"Input file not found: {inputFilePath}");

            // Load the workbook from the specified file.
            Workbook workbook = new Workbook(inputFilePath);

            // Work with the first worksheet (adjust as needed).
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Determine the last used row in the source column.
            int lastRow = cells.MaxDataRow; // zero‑based index of the last row that contains data in any column.

            // Define the index for the running total column (right after the source column).
            int totalColumnIndex = sourceColumnIndex + 1;

            // Insert a header for the running total column if a header row exists.
            if (headerRowIndex >= 0)
            {
                cells[headerRowIndex, totalColumnIndex].PutValue("Running Total");
            }

            // Starting row for data (skip header if present).
            int dataStartRow = headerRowIndex >= 0 ? headerRowIndex + 1 : 0;

            // Helper to convert column index to Excel column letter (e.g., 0 -> "A").
            string GetColumnLetter(int colIndex)
            {
                int dividend = colIndex + 1;
                string columnName = string.Empty;
                while (dividend > 0)
                {
                    int modulo = (dividend - 1) % 26;
                    columnName = Convert.ToChar(65 + modulo) + columnName;
                    dividend = (dividend - modulo) / 26;
                }
                return columnName;
            }

            string sourceColumnLetter = GetColumnLetter(sourceColumnIndex);

            // Populate each cell in the running total column with a SUM formula that accumulates values from the first data row down to the current row.
            for (int row = dataStartRow; row <= lastRow; row++)
            {
                // Build the range string, e.g., "A2:A5".
                string range = $"{sourceColumnLetter}{dataStartRow + 1}:{sourceColumnLetter}{row + 1}";
                // Set the SUM formula.
                cells[row, totalColumnIndex].Formula = $"=SUM({range})";
            }

            // Ensure the output directory exists.
            string outputDir = Path.GetDirectoryName(outputFilePath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook to the output path.
            workbook.Save(outputFilePath);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            // Example usage with default file names.
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                RunningTotalInserter.InsertRunningTotal(inputPath, outputPath, sourceColumnIndex: 0, headerRowIndex: 0);
                Console.WriteLine($"Running total inserted successfully. Output saved to '{outputPath}'.");
            }
            catch (FileNotFoundException ex)
            {
                Console.Error.WriteLine($"File error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
