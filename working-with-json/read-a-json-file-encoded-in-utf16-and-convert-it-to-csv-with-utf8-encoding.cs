// Title: Read a UTF‑16 JSON file and export it as a UTF‑8 CSV using Aspose.Cells in C#
// Common Searches: c# read utf-16 json file and convert to csv with aspose.cells | aspose.cells export json array to csv with utf-8 encoding | how to use TxtSaveOptions to save csv in utf-8 from json data in c# | convert json objects to spreadsheet rows and save as csv using aspose.cells | parse unicode json and generate utf-8 csv file in .net
// Tags: aspose.cells json to csv conversion | utf-16 json parsing c# | utf-8 csv export aspose.cells | json array to workbook rows | txtsaveoptions encoding csv | c# workbook save as csv with custom encoding

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace JsonToCsvConverter
{
    // The sample reads a UTF‑16 encoded JSON file containing an array of objects, extracts column headers from the first object, populates an Aspose.Cells Workbook with the data, and saves the workbook as a CSV file encoded in UTF‑8 using TxtSaveOptions. It also creates the output directory if needed and includes basic validation and error handling.
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: JsonToCsvConverter <inputJsonUtf16Path> <outputCsvUtf8Path>");
                return;
            }

            string inputJsonPath = args[0];
            string outputCsvPath = args[1];

            // Verify that the input file exists before proceeding.
            if (!File.Exists(inputJsonPath))
            {
                Console.WriteLine($"Error: Input JSON file not found: {inputJsonPath}");
                return;
            }

            try
            {
                // 1. Read the JSON file using UTF‑16 (Unicode) encoding.
                string jsonText = File.ReadAllText(inputJsonPath, Encoding.Unicode);

                // 2. Parse the JSON. Expecting an array of objects (each object = one row).
                using JsonDocument document = JsonDocument.Parse(jsonText);
                JsonElement root = document.RootElement;

                if (root.ValueKind != JsonValueKind.Array)
                {
                    Console.WriteLine("The JSON root element must be an array of objects.");
                    return;
                }

                // 3. Extract column names from the first object and collect rows.
                List<string> headers = new List<string>();
                List<Dictionary<string, string>> rows = new List<Dictionary<string, string>>();
                bool firstRow = true;

                foreach (JsonElement element in root.EnumerateArray())
                {
                    if (element.ValueKind != JsonValueKind.Object)
                    {
                        Console.WriteLine("Each array element must be a JSON object.");
                        return;
                    }

                    var row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    foreach (JsonProperty prop in element.EnumerateObject())
                    {
                        string propName = prop.Name;
                        string propValue = prop.Value.ToString();

                        row[propName] = propValue;

                        if (firstRow && !headers.Contains(propName))
                            headers.Add(propName);
                    }

                    rows.Add(row);
                    firstRow = false;
                }

                // 4. Create a workbook and populate it with the extracted data.
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];
                var cells = sheet.Cells;

                // Write header row.
                for (int col = 0; col < headers.Count; col++)
                    cells[0, col].PutValue(headers[col]);

                // Write data rows.
                for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
                {
                    var row = rows[rowIndex];
                    for (int col = 0; col < headers.Count; col++)
                    {
                        string header = headers[col];
                        if (row.TryGetValue(header, out string value))
                            cells[rowIndex + 1, col].PutValue(value);
                        else
                            cells[rowIndex + 1, col].PutValue(string.Empty);
                    }
                }

                // 5. Save the workbook as CSV with UTF‑8 encoding.
                var csvOptions = new TxtSaveOptions(SaveFormat.Csv)
                {
                    Encoding = Encoding.UTF8,
                    Separator = ',',
                    CreateDirectory = true // ensure output directory is created if missing
                };

                // Ensure the output directory exists.
                string outputDir = Path.GetDirectoryName(outputCsvPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    Directory.CreateDirectory(outputDir);

                workbook.Save(outputCsvPath, csvOptions);

                Console.WriteLine($"Successfully converted '{inputJsonPath}' to UTF‑8 CSV at '{outputCsvPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
