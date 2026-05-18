using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonFlattenExample
{
    public class JsonFlattenProcessor
    {
        public static void Run()
        {
            try
            {
                // Input JSON file containing nested objects
                string inputJsonPath = "nested_input.json";

                // Output JSON file that will contain the flattened structure
                string outputJsonPath = "flattened_output.json";

                // Ensure the input file exists
                if (!File.Exists(inputJsonPath))
                {
                    Console.WriteLine($"Input file not found: {inputJsonPath}");
                    return;
                }

                // Read the JSON content
                string jsonContent = File.ReadAllText(inputJsonPath);

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Import the JSON data into the worksheet using default layout options
                JsonUtility.ImportData(jsonContent, worksheet.Cells, 0, 0, new JsonLayoutOptions());

                // Configure JSON save options for a flattened structure
                JsonSaveOptions saveOptions = new JsonSaveOptions
                {
                    ExportNestedStructure = false,      // Flatten the JSON
                    AlwaysExportAsJsonObject = true,    // Ensure output is a JSON object even with one sheet
                    SkipEmptyRows = true,               // Omit empty rows for cleaner output
                    HasHeaderRow = true,                // Treat first row as header (optional)
                    ExportEmptyCells = false            // Do not include empty cells
                };

                // Save the workbook as a JSON file with the flattening options applied
                workbook.Save(outputJsonPath, saveOptions);

                // Display the resulting flattened JSON
                if (File.Exists(outputJsonPath))
                {
                    string flattenedJson = File.ReadAllText(outputJsonPath);
                    Console.WriteLine("Flattened JSON output:");
                    Console.WriteLine(flattenedJson);
                }
                else
                {
                    Console.WriteLine($"Failed to create output file: {outputJsonPath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            JsonFlattenProcessor.Run();
        }
    }
}