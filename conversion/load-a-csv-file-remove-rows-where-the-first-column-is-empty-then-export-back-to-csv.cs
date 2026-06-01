using System;
using System.IO;
using Aspose.Cells;

namespace CsvProcessingExample
{
    class Program
    {
        static void Main()
        {
            // Paths for input and output CSV files
            string inputCsvPath = "input.csv";
            string outputCsvPath = "output.csv";

            try
            {
                // Verify that the input CSV file exists
                if (!File.Exists(inputCsvPath))
                {
                    Console.WriteLine($"Input file not found: {inputCsvPath}");
                    return;
                }

                // Create a new workbook instance
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Load CSV data into the worksheet using TxtLoadOptions
                TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv)
                {
                    Separator = ',',                     // Define the column separator
                    TreatConsecutiveDelimitersAsOne = false,
                    ConvertNumericData = true
                };
                cells.ImportCSV(inputCsvPath, loadOptions, 0, 0);

                // Iterate rows in reverse order and delete rows where the first column is empty
                for (int row = cells.MaxDataRow; row >= 0; row--)
                {
                    Cell firstCell = cells[row, 0];

                    // Determine if the first cell is blank or contains only whitespace
                    bool isEmpty = firstCell.Type == CellValueType.IsNull ||
                                   string.IsNullOrWhiteSpace(firstCell.StringValue);

                    if (isEmpty)
                    {
                        // Delete the entire row
                        cells.DeleteRow(row);
                    }
                }

                // Save the modified worksheet back to CSV
                TxtSaveOptions saveOptions = new TxtSaveOptions
                {
                    Separator = ',',                     // Use comma as separator for CSV
                    TrimLeadingBlankRowAndColumn = true // Trim any leading blank rows/columns
                };
                workbook.Save(outputCsvPath, saveOptions);

                Console.WriteLine("Processing completed. Output saved to: " + outputCsvPath);
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}