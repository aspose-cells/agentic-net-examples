// Title: Import CSV file with quoted fields into an Aspose.Cells workbook using a custom ICellsDataTable in C#
// AI Prompts: Write a C# console program that reads a CSV file, parses commas and double‑quoted fields according to RFC 4180, builds a two‑dimensional object array, creates an ICellsDataTable via CellsDataTableFactory, and imports the data into a worksheet. | Show how to set ImportTableOptions to automatically convert numeric strings and include the header row when importing a custom ICellsDataTable into an Excel workbook with Aspose.Cells.
// Common Searches: c# aspose.cells import csv with quoted fields using icellsdatatable | how to parse RFC 4180 csv in C# before importing to Aspose.Cells | using CellsDataTableFactory to load object array into an Excel workbook | aspose.cells ImportTableOptions convert numeric strings from csv | read csv file and import to Aspose.Cells worksheet with header row
// Tags: ICellsDataTable CSV import C# | CellsDataTableFactory object array import | ImportData ImportTableOptions numeric conversion | RFC 4180 CSV parser Aspose.Cells | Aspose.Cells worksheet import custom data source

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvImport
{
    // The example reads a CSV file, parses each line handling commas and double‑quoted fields per RFC 4180, builds a two‑dimensional object array with a header row, creates an ICellsDataTable via CellsDataTableFactory, imports it into the first worksheet of a new Workbook using ImportTableOptions (including numeric conversion and header display), and saves the workbook as an XLSX file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the CSV file (replace with your actual file path)
                string csvPath = "sample.csv";

                // Verify that the CSV file exists before attempting to read it
                if (!File.Exists(csvPath))
                {
                    Console.WriteLine($"CSV file not found: {Path.GetFullPath(csvPath)}");
                    return;
                }

                // Read all lines from the CSV file
                string[] lines = File.ReadAllLines(csvPath);

                if (lines.Length == 0)
                {
                    Console.WriteLine("CSV file is empty.");
                    return;
                }

                // Parse each line into a list of fields handling commas and quotes
                List<string[]> parsedRows = new List<string[]>();
                foreach (string line in lines)
                {
                    parsedRows.Add(ParseCsvLine(line));
                }

                // First row is assumed to be the header
                string[] header = parsedRows[0];
                // Remaining rows are data
                List<string[]> dataRows = parsedRows.GetRange(1, parsedRows.Count - 1);

                // Convert data rows to object[][] required by CellsDataTableFactory
                object[][] dataObjects = new object[dataRows.Count][];
                for (int i = 0; i < dataRows.Count; i++)
                {
                    // Each field is kept as string; conversion to numeric/date will be handled by ImportTableOptions
                    dataObjects[i] = dataRows[i];
                }

                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Obtain the factory from the workbook
                CellsDataTableFactory factory = workbook.CellsDataTableFactory;

                // Create ICellsDataTable from the parsed data
                // hasHeader = true indicates that the first row (header) is not part of the data rows
                ICellsDataTable dataTable = factory.GetInstance(dataObjects, true, header);

                // Set import options (e.g., convert numeric strings to numbers)
                ImportTableOptions importOptions = new ImportTableOptions
                {
                    ConvertNumericData = true,
                    IsFieldNameShown = true // import header as first row in the sheet
                };

                // Import the custom data table into the worksheet starting at cell A1 (row 0, column 0)
                cells.ImportData(dataTable, 0, 0, importOptions);

                // Save the workbook
                string outputPath = "ImportedFromCsv.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"CSV data imported successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Parses a single CSV line handling commas and double quotes according to RFC 4180
        private static string[] ParseCsvLine(string line)
        {
            List<string> fields = new List<string>();
            bool insideQuotes = false;
            StringBuilder field = new StringBuilder();

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (c == '\"')
                {
                    // If double quote inside quoted field and next char is also a quote, treat as escaped quote
                    if (insideQuotes && i + 1 < line.Length && line[i + 1] == '\"')
                    {
                        field.Append('\"');
                        i++; // skip the escaped quote
                    }
                    else
                    {
                        // Toggle the insideQuotes flag
                        insideQuotes = !insideQuotes;
                    }
                }
                else if (c == ',' && !insideQuotes)
                {
                    // End of field
                    fields.Add(field.ToString());
                    field.Clear();
                }
                else
                {
                    field.Append(c);
                }
            }

            // Add the last field
            fields.Add(field.ToString());

            return fields.ToArray();
        }
    }
}
