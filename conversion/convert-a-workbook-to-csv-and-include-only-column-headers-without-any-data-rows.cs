using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsHeaderOnlyCsv
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output_headers_only.csv";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(inputPath);
                Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

                // Determine the last used column in the first row (header row)
                // Use MaxColumn which returns the zero‑based index of the last column that contains data
                int lastHeaderColumn = sourceSheet.Cells.MaxColumn;

                // Create a new workbook that will contain only the header row
                Workbook headerOnlyWorkbook = new Workbook();
                Worksheet headerSheet = headerOnlyWorkbook.Worksheets[0];

                // Copy header values from the source sheet to the new workbook
                for (int col = 0; col <= lastHeaderColumn; col++)
                {
                    string headerValue = sourceSheet.Cells[0, col].StringValue;
                    headerSheet.Cells[0, col].PutValue(headerValue);
                }

                // Prepare CSV save options
                TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
                {
                    ExportAllSheets = false,
                    TrimLeadingBlankRowAndColumn = true
                };

                // Save the header‑only workbook as a CSV file
                headerOnlyWorkbook.Save(outputPath, csvOptions);
                Console.WriteLine($"Header‑only CSV saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}