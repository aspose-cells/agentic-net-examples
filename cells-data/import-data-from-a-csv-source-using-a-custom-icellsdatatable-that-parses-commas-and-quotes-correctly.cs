// Title: Import CSV with quoted fields into Aspose.Cells using a custom ICellsDataTable (C#)
// Description: Reads a CSV file, parses commas and double‑quoted values with a custom method, builds an ICellsDataTable via CellsDataTableFactory, imports the table into a worksheet using ImportData, and saves the result as an XLSX workbook.
// Keywords: Aspose.Cells CSV import | ICellsDataTable | CellsDataTableFactory | ParseCsvLine C# | ImportData ImportTableOptions | C# CSV parsing quotes | Excel workbook from CSV | Aspose.Cells .NET
// Common Searches: Aspose.Cells import CSV with quotes | Create ICellsDataTable from list of rows | C# parse CSV line with double quotes | ImportData example Aspose.Cells | Handle commas inside CSV fields using Aspose.Cells
// Developer Intent: Read a CSV file, convert each line to a row collection, build an ICellsDataTable, and import it into an Excel worksheet.
// Use Cases: Load a sales‑report CSV where product names contain commas or quotes, convert numeric strings, and generate an analysis workbook. | Transform a configuration CSV with mixed data types into an XLSX file for distribution to non‑technical stakeholders. | Process a small in‑memory CSV, apply custom parsing rules, and preserve original formatting when creating a formatted report workbook.
// AI Prompts: Show how to stream a large CSV into an ICellsDataTable without loading the entire file into memory. | Provide code to set column widths and header styles via ImportTableOptions while importing CSV data. | Explain how to modify ParseCsvLine to support a semicolon delimiter and still use CellsDataTableFactory.

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCsvImportExample
{
    // Reads a CSV file, parses commas and double‑quoted values with a custom method, builds an ICellsDataTable via CellsDataTableFactory, imports the table into a worksheet using ImportData, and saves the result as an XLSX workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the CSV file (replace with your actual file path)
                string csvPath = "sample.csv";

                // Verify that the CSV file exists to avoid FileNotFoundException
                if (!File.Exists(csvPath))
                {
                    Console.WriteLine($"CSV file not found: {Path.GetFullPath(csvPath)}");
                    return;
                }

                // Read all lines from the CSV file
                string[] lines = File.ReadAllLines(csvPath);

                // Prepare a collection of rows, each row is an ArrayList of objects
                ArrayList dataLists = new ArrayList();

                foreach (string line in lines)
                {
                    // Parse a single CSV line handling commas and double quotes
                    ArrayList fields = ParseCsvLine(line);
                    dataLists.Add(fields);
                }

                // Create a new workbook (creation rule)
                Workbook workbook = new Workbook();

                // Obtain a CellsDataTableFactory from the workbook
                CellsDataTableFactory factory = workbook.CellsDataTableFactory;

                // Build an ICellsDataTable from the parsed data (custom data table rule)
                ICellsDataTable dataTable = factory.GetInstance(dataLists, true);

                // Import the data table into the first worksheet starting at cell A1
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet.Cells.ImportData(dataTable, 0, 0, new ImportTableOptions());

                // Save the workbook (save rule)
                string outputPath = "ImportedFromCsv.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Parses a CSV line respecting commas and double‑quoted fields.
        // Returns an ArrayList where each element is a string (or numeric if convertible).
        private static ArrayList ParseCsvLine(string line)
        {
            ArrayList result = new ArrayList();
            int i = 0;
            while (i < line.Length)
            {
                // Skip leading whitespace (optional)
                while (i < line.Length && char.IsWhiteSpace(line[i]))
                    i++;

                string field;
                if (i < line.Length && line[i] == '\"')
                {
                    // Quoted field
                    i++; // skip opening quote
                    int start = i;
                    while (i < line.Length)
                    {
                        // Look for closing quote; double quotes inside are escaped by another quote
                        if (line[i] == '\"')
                        {
                            if (i + 1 < line.Length && line[i + 1] == '\"')
                            {
                                // Escaped quote, skip one and continue
                                i += 2;
                            }
                            else
                            {
                                // End of quoted field
                                break;
                            }
                        }
                        else
                        {
                            i++;
                        }
                    }
                    field = line.Substring(start, i - start).Replace("\"\"", "\"");
                    i++; // skip closing quote
                }
                else
                {
                    // Unquoted field
                    int start = i;
                    while (i < line.Length && line[i] != ',')
                        i++;
                    field = line.Substring(start, i - start);
                }

                // Add the field (attempt numeric conversion)
                if (double.TryParse(field, out double num))
                    result.Add(num);
                else
                    result.Add(field);

                // Skip delimiter
                if (i < line.Length && line[i] == ',')
                    i++;
            }

            return result;
        }
    }
}
