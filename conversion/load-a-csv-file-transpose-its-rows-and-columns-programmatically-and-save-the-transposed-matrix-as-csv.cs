using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTransposeCsv
{
    class Program
    {
        static void Main()
        {
            // Input and output CSV file paths
            string inputCsvPath = "input.csv";
            string outputCsvPath = "transposed.csv";

            try
            {
                // Verify that the input CSV file exists
                if (!File.Exists(inputCsvPath))
                {
                    Console.WriteLine($"Input file not found: {inputCsvPath}");
                    return;
                }

                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Import the CSV data into the worksheet (lifecycle rule: load)
                // Parameters: file name, delimiter, convert numeric data, start row, start column
                cells.ImportCSV(inputCsvPath, ",", true, 0, 0);

                // Get the used range after import
                Aspose.Cells.Range usedRange = cells.MaxDisplayRange;

                // Transpose the used range (rotate rows to columns)
                usedRange.Transpose();

                // Save the transposed data as a CSV file (lifecycle rule: save)
                workbook.Save(outputCsvPath, SaveFormat.Csv);

                Console.WriteLine($"Transposed CSV saved to: {outputCsvPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}