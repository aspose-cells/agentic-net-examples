// Title: C# – Load CSV, Delete Rows with Empty First Column, and Export Clean CSV using Aspose.Cells
// Description: A C# example that loads a CSV into an Aspose.Cells workbook, removes every row whose first column is blank, and saves the cleaned data back to CSV with custom separator and trimmed leading blanks.
// Keywords: Aspose.Cells CSV import C# | delete rows with empty first column | remove blank rows from CSV | export CSV Aspose.Cells | TxtSaveOptions separator | trim leading blank rows | CSV cleaning .NET | Aspose.Cells workbook to CSV
// Common Searches: how to delete rows with empty column A using Aspose.Cells | Aspose.Cells C# save worksheet as CSV with custom separator | remove empty rows from CSV programmatically .NET | trim leading blank rows when exporting CSV Aspose.Cells | load CSV, filter rows, and export with Aspose.Cells
// Developer Intent: Filter out rows whose first column is empty from a CSV file and write the resulting dataset back to a new CSV using Aspose.Cells.
// Use Cases: Clean raw data sets by discarding records missing a primary key before analytics. | Prepare bulk‑import CSV files for ERP or CRM systems, ensuring no leading‑column gaps. | Generate concise CSV reports after eliminating incomplete rows to reduce downstream errors.
// AI Prompts: Generate C# code with Aspose.Cells that reads a CSV, removes rows where column A is empty, and writes the result to a new CSV. | Show how to configure TxtSaveOptions in Aspose.Cells to trim leading blank rows and set a comma separator when exporting to CSV.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCsvProcessing
{
    // A C# example that loads a CSV into an Aspose.Cells workbook, removes every row whose first column is blank, and saves the cleaned data back to CSV with custom separator and trimmed leading blanks.
    class Program
    {
        static void Main()
        {
            // Input and output CSV file paths
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

                // Create a new workbook and import the CSV data into the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Import CSV (comma‑separated, convert numeric data, start at A1)
                cells.ImportCSV(inputCsvPath, ",", true, 0, 0);

                // Iterate rows from bottom to top and delete rows whose first column is empty
                for (int row = cells.MaxDataRow; row >= 0; row--)
                {
                    // Get the cell in the first column (index 0)
                    Cell firstCell = cells[row, 0];

                    // Consider a cell empty if it has no value or its string representation is empty
                    bool isEmpty = firstCell.Type == CellValueType.IsNull ||
                                   string.IsNullOrEmpty(firstCell.StringValue);

                    if (isEmpty)
                    {
                        // Delete the entire row
                        cells.DeleteRow(row);
                    }
                }

                // Prepare CSV save options (trim leading blanks and set separator)
                TxtSaveOptions saveOptions = new TxtSaveOptions
                {
                    TrimLeadingBlankRowAndColumn = true,
                    Separator = ','
                };

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputCsvPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook back to CSV
                workbook.Save(outputCsvPath, saveOptions);

                Console.WriteLine("Processing completed. Output saved to: " + outputCsvPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during processing:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
