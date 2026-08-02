// Title: C# – Transpose CSV rows and columns using Aspose.Cells and save as new CSV
// Description: Loads a CSV into an Aspose.Cells workbook, applies Range.Transpose to swap rows with columns, and writes the transposed matrix back to a CSV file. The sample validates the input file and handles runtime errors.
// Keywords: Aspose.Cells CSV transpose C# | Range.Transpose Aspose.Cells | ImportCSV to workbook | SaveFormat.Csv Aspose | C# transpose rows to columns | .NET CSV data rotation | Excel library CSV manipulation | programmatic CSV matrix flip
// Common Searches: how to transpose a csv with Aspose.Cells in C# | c# transpose rows and columns csv file | aspocells importcsv and transpose example | save transposed worksheet as csv using Aspose | range.transpose method usage .net
// Developer Intent: Read an existing CSV, exchange its rows and columns programmatically, and output the result as a new CSV using the Aspose.Cells for .NET API.
// Use Cases: Re‑orienting data for reporting tools that expect the opposite axis layout. | Preparing matrix‑style CSVs for legacy systems that require rows as columns. | Automating a lightweight ETL step where CSV orientation must be flipped without Excel UI.
// AI Prompts: Write C# code that uses Aspose.Cells to load a CSV, transpose the data, and save it as another CSV. | Explain the limitations of Range.Transpose in Aspose.Cells, such as maximum worksheet size and memory impact. | Create robust error handling for the CSV transposition sample, covering missing files, empty inputs, and unsupported delimiters.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTransposeCsv
{
    // Loads a CSV into an Aspose.Cells workbook, applies Range.Transpose to swap rows with columns, and writes the transposed matrix back to a CSV file. The sample validates the input file and handles runtime errors.
    class Program
    {
        static void Main()
        {
            // Paths for input and output CSV files
            string inputCsv = "input.csv";
            string outputCsv = "transposed.csv";

            try
            {
                // Verify that the input CSV file exists
                if (!File.Exists(inputCsv))
                {
                    Console.WriteLine($"Error: Input file '{inputCsv}' not found.");
                    return;
                }

                // 1. Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // 2. Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // 3. Import the CSV data starting at cell A1 (row 0, column 0)
                //    Using comma as delimiter and converting numeric data
                cells.ImportCSV(inputCsv, ",", true, 0, 0);

                // 4. Determine the size of the imported data
                //    MaxDataRow/MaxDataColumn give zero‑based indexes of the last used cell
                int lastRow = cells.MaxDataRow;      // zero‑based
                int lastCol = cells.MaxDataColumn;   // zero‑based

                // 5. Create a range that covers the entire imported area
                //    CreateRange(firstRow, firstColumn, totalRows, totalColumns)
                Aspose.Cells.Range dataRange = cells.CreateRange(0, 0, lastRow + 1, lastCol + 1);

                // 6. Transpose the range (rotate rows ↔ columns)
                dataRange.Transpose();

                // 7. Save the workbook as a CSV file (lifecycle rule: save)
                workbook.Save(outputCsv, SaveFormat.Csv);

                Console.WriteLine($"CSV file has been transposed and saved to '{outputCsv}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
