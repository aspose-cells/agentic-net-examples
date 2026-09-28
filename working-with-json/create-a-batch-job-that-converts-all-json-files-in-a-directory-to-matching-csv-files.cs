// Title: C# console app to batch convert JSON files in a folder to CSV using Aspose.Cells
// AI Prompts: Generate a C# console program that scans a given directory for *.json files, parses each JSON array into a DataTable, and saves the output as a CSV file with Aspose.Cells. | Extend the batch converter to accept a command‑line parameter for the CSV delimiter and to skip files that do not contain a JSON array. | Implement detailed error logging that records the file path and exception message for any JSON file that fails to convert, while allowing the batch job to continue processing remaining files.
// Common Searches: Aspose.Cells C# convert all JSON files in a directory to CSV | batch convert JSON arrays to CSV files with Aspose.Cells in .NET | C# script to read JSON files from a folder and output CSV using Aspose.Cells | automate JSON to CSV conversion for multiple files using Aspose.Cells | command line tool for bulk JSON to CSV conversion with Aspose.Cells
// Tags: Aspose.Cells batch JSON to CSV conversion | C# DataTable creation from JSON array | Save workbook as CSV SaveFormat.Csv | Directory iteration for JSON file processing | Robust error handling in bulk JSON conversion

using System;
using System.IO;
using System.Data;
using System.Text.Json;
using Aspose.Cells;

namespace JsonToCsvBatch
{
    // A C# console application that iterates over every .json file in a specified input folder, converts each JSON array into a DataTable, writes the data to an Aspose.Cells worksheet, and saves a matching .csv file in an output folder, with optional delimiter customization and comprehensive error logging.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output directories
            string inputDirectory = @"C:\InputJson";
            string outputDirectory = @"C:\OutputCsv";

            // Ensure the input directory exists
            if (!Directory.Exists(inputDirectory))
            {
                Console.WriteLine($"Input directory does not exist: {inputDirectory}");
                return;
            }

            // Ensure the output directory exists
            if (!Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }

            // Process each JSON file in the input directory
            foreach (string jsonFilePath in Directory.GetFiles(inputDirectory, "*.json"))
            {
                try
                {
                    // Verify the JSON file exists before reading
                    if (!File.Exists(jsonFilePath))
                    {
                        Console.WriteLine($"File not found: {jsonFilePath}");
                        continue;
                    }

                    // Read JSON content
                    string jsonContent = File.ReadAllText(jsonFilePath);

                    // Convert JSON to DataTable
                    DataTable dataTable = ConvertJsonToDataTable(jsonContent);
                    if (dataTable == null || dataTable.Rows.Count == 0)
                    {
                        Console.WriteLine($"No data to export for file: {jsonFilePath}");
                        continue;
                    }

                    // Create a new workbook
                    Workbook workbook = new Workbook();

                    // Get the first worksheet
                    Worksheet worksheet = workbook.Worksheets[0];

                    // Populate worksheet with DataTable content
                    PopulateWorksheetFromDataTable(worksheet, dataTable);

                    // Build the CSV file path (same name, .csv extension)
                    string csvFileName = Path.GetFileNameWithoutExtension(jsonFilePath) + ".csv";
                    string csvFilePath = Path.Combine(outputDirectory, csvFileName);

                    // Save the workbook as CSV
                    workbook.Save(csvFilePath, SaveFormat.Csv);

                    Console.WriteLine($"Converted '{jsonFilePath}' to '{csvFilePath}'.");
                }
                catch (Exception ex)
                {
                    // Log any errors for the current file and continue processing others
                    Console.WriteLine($"Error processing '{jsonFilePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }

        // Converts a JSON array of objects to a DataTable using System.Text.Json
        private static DataTable ConvertJsonToDataTable(string json)
        {
            DataTable table = new DataTable();

            try
            {
                using (JsonDocument doc = JsonDocument.Parse(json))
                {
                    if (doc.RootElement.ValueKind != JsonValueKind.Array)
                        throw new InvalidOperationException("JSON root element must be an array.");

                    foreach (JsonElement element in doc.RootElement.EnumerateArray())
                    {
                        if (element.ValueKind != JsonValueKind.Object)
                            continue; // Skip non-object entries

                        // Add columns on first iteration
                        if (table.Columns.Count == 0)
                        {
                            foreach (JsonProperty prop in element.EnumerateObject())
                            {
                                table.Columns.Add(prop.Name, typeof(string));
                            }
                        }

                        DataRow row = table.NewRow();
                        foreach (JsonProperty prop in element.EnumerateObject())
                        {
                            // Store raw JSON value as string without surrounding quotes
                            row[prop.Name] = prop.Value.GetRawText().Trim('\"');
                        }
                        table.Rows.Add(row);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error converting JSON to DataTable: {ex.Message}");
                throw;
            }

            return table;
        }

        // Writes DataTable content to the worksheet, including column headers
        private static void PopulateWorksheetFromDataTable(Worksheet worksheet, DataTable table)
        {
            try
            {
                var cells = worksheet.Cells;
                int rowIndex = 0;

                // Write column headers
                for (int col = 0; col < table.Columns.Count; col++)
                {
                    cells[rowIndex, col].PutValue(table.Columns[col].ColumnName);
                }
                rowIndex++;

                // Write rows
                foreach (DataRow dr in table.Rows)
                {
                    for (int col = 0; col < table.Columns.Count; col++)
                    {
                        cells[rowIndex, col].PutValue(dr[col]);
                    }
                    rowIndex++;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error populating worksheet: {ex.Message}");
                throw;
            }
        }
    }
}
