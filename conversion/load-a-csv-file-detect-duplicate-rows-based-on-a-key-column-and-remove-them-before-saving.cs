// Title: C# – Remove Duplicate Rows from CSV and Save as XLSX with Aspose.Cells
// Description: Load a CSV file into an Aspose.Cells workbook, use Cells.RemoveDuplicates to drop rows that share the same key column (keeping the header), and export the cleaned data to an XLSX file.
// Keywords: Aspose.Cells | C# | CSV duplicate removal | RemoveDuplicates | cells.RemoveDuplicates example | convert CSV to XLSX | data cleaning with Aspose.Cells | Excel export C# | Aspose.Cells API
// Common Searches: Aspose.Cells remove duplicate rows from CSV | C# import CSV and delete duplicate records | How to use Cells.RemoveDuplicates in .NET | Convert CSV to Excel without duplicates Aspose | Aspose.Cells CSV to XLSX duplicate elimination
// Developer Intent: Import a CSV, eliminate rows with duplicate key values, and save the result as an Excel workbook.
// Use Cases: Prepare a clean dataset for BI tools by removing duplicate identifiers before conversion to Excel. | De‑duplicate log or export files programmatically and generate a shareable XLSX report. | Automate data‑quality checks in ETL pipelines where CSV sources must contain unique records.
// AI Prompts: Write C# code that uses Aspose.Cells to read a CSV, remove duplicate rows based on a specified column, and save as XLSX. | Explain each argument of the Cells.RemoveDuplicates method shown in the example. | Suggest robust error handling and logging for missing CSV files and duplicate‑removal failures in the Aspose.Cells workflow.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Load a CSV file into an Aspose.Cells workbook, use Cells.RemoveDuplicates to drop rows that share the same key column (keeping the header), and export the cleaned data to an XLSX file.
    public class RemoveDuplicateRowsFromCsv
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Path to the source CSV file
            string csvPath = "input.csv";

            // Verify that the CSV file exists to avoid FileNotFoundException
            if (!File.Exists(csvPath))
            {
                Console.WriteLine($"CSV file not found: {csvPath}");
                return;
            }

            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Import the CSV data starting at cell A1 (row 0, column 0)
                // Using comma as delimiter and converting numeric strings to numbers
                cells.ImportCSV(csvPath, ",", true, 0, 0);

                // Determine the used range after import
                int startRow = 0;                         // First row (including header)
                int startColumn = 0;                      // First column (key column)
                int endRow = cells.MaxDataRow;            // Last row with data
                int endColumn = cells.MaxDataColumn;      // Last column with data

                // Remove duplicate rows based on the first column (index 0)
                // Assume the first row contains headers (hasHeaders = true)
                // columnOffsets specifies which columns are used for duplicate comparison
                cells.RemoveDuplicates(startRow, startColumn, endRow, endColumn, true, new int[] { 0 });

                // Save the cleaned workbook to an XLSX file
                string outputPath = "output.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Processed file saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing CSV: {ex.Message}");
            }
        }
    }
}
