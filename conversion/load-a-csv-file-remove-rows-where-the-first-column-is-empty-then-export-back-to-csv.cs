// Title: C# – Remove rows with empty first column from CSV using Aspose.Cells and export
// Description: Loads a CSV into an Aspose.Cells workbook, deletes every row whose first column is blank or whitespace, and saves the cleaned worksheet back to CSV with TrimLeadingBlankRowAndColumn enabled.
// Keywords: Aspose.Cells CSV import C# | delete empty rows Aspose.Cells | remove rows with blank first column | save worksheet to CSV .NET | TxtSaveOptions TrimLeadingBlankRowAndColumn | C# CSV cleanup Aspose | filter CSV rows Aspose.Cells
// Common Searches: how to delete rows with empty column A using Aspose.Cells | C# Aspose.Cells import CSV and remove blank rows | save cleaned CSV with Aspose.Cells TxtSaveOptions | remove empty rows from CSV in .NET | Aspose.Cells filter rows by first column
// Developer Intent: Load a CSV, drop rows where column A is empty, and write the result to a new CSV file.
// Use Cases: Pre‑process exported data before loading into analytics tools that reject blank leading columns. | Clean log files or data extracts automatically in batch jobs. | Prepare CSV inputs for legacy systems that cannot handle rows with missing key values.
// AI Prompts: Write C# code that uses Aspose.Cells to read a CSV, remove rows with an empty first column, and save the output as a new CSV. | Explain the impact of TxtSaveOptions.TrimLeadingBlankRowAndColumn when exporting a worksheet to CSV. | Suggest a LINQ‑based method to filter rows by the first column without iterating backwards.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCsvProcessing
{
    // Loads a CSV into an Aspose.Cells workbook, deletes every row whose first column is blank or whitespace, and saves the cleaned worksheet back to CSV with TrimLeadingBlankRowAndColumn enabled.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source CSV file
                string inputCsv = "input.csv";

                // Verify that the input CSV file exists
                if (!File.Exists(inputCsv))
                {
                    Console.WriteLine($"Error: Input file \"{inputCsv}\" not found.");
                    return;
                }

                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Import the CSV file into the worksheet (starting at cell A1)
                // Using comma as separator, converting numeric data, and starting at row 0, column 0
                cells.ImportCSV(inputCsv, ",", true, 0, 0);

                // Iterate from the last data row upwards and delete rows where the first column is empty
                for (int row = cells.MaxDataRow; row >= 0; row--)
                {
                    // Get the string value of the first column (column index 0)
                    string firstColValue = cells[row, 0].StringValue;

                    // If the cell is empty or contains only whitespace, delete the entire row
                    if (string.IsNullOrWhiteSpace(firstColValue))
                    {
                        // Delete the row using the Rows collection
                        cells.Rows.RemoveAt(row);
                    }
                }

                // Prepare CSV save options (trim leading blank rows/columns as Excel does)
                TxtSaveOptions saveOptions = new TxtSaveOptions
                {
                    TrimLeadingBlankRowAndColumn = true,
                    Separator = ','
                };

                // Save the modified worksheet back to a CSV file
                string outputCsv = "output.csv";
                workbook.Save(outputCsv, saveOptions);

                Console.WriteLine("Processing completed. Output saved to " + outputCsv);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
