using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonErrorHandling
{
    class Program
    {
        static void Main()
        {
            // Path to the source JSON file (may be malformed)
            string jsonFilePath = "data.json";

            // Ensure the file exists for demonstration purposes
            if (!File.Exists(jsonFilePath))
            {
                // Create a deliberately malformed JSON file
                File.WriteAllText(jsonFilePath, "{ \"Name\": \"John\", \"Age\": 30, "); // missing closing brace
            }

            // Create JSON load options (optional configuration)
            JsonLoadOptions loadOptions = new JsonLoadOptions
            {
                // KeepSchema property does not exist; using default options
                StartCell = "A1"
            };

            try
            {
                // Verify the JSON file exists before attempting to load it
                if (!File.Exists(jsonFilePath))
                {
                    Console.WriteLine($"File not found: {jsonFilePath}");
                    return;
                }

                // Load the JSON file into a workbook using the specified options
                Workbook workbook = new Workbook(jsonFilePath, loadOptions);
                workbook.Save("output.xlsx");
                Console.WriteLine("JSON loaded and workbook saved successfully.");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.FileCorrupted ||
                                            ex.Code == ExceptionType.InvalidData ||
                                            ex.Code == ExceptionType.IO)
            {
                // Specific handling for JSON parsing related errors.
                Console.WriteLine($"Failed to load JSON file: {ex.Message}");
                Console.WriteLine($"Error Code: {ex.Code}");
            }
            catch (Exception ex)
            {
                // General fallback for any other unexpected exceptions.
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }

            // Alternative approach using JsonUtility.ImportData with a JSON string.
            string malformedJson = "[{ \"Name\": \"Alice\", \"Age\": 25 }, { \"Name\": \"Bob\", \"Age\": } ]"; // Age value missing

            try
            {
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];

                // Import the JSON string into the worksheet.
                JsonUtility.ImportData(malformedJson, ws.Cells, 0, 0, new JsonLayoutOptions());

                // Save if import succeeds.
                wb.Save("imported_output.xlsx");
                Console.WriteLine("JSON string imported and workbook saved successfully.");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.FileCorrupted ||
                                            ex.Code == ExceptionType.InvalidData ||
                                            ex.Code == ExceptionType.IO)
            {
                Console.WriteLine($"Failed to import JSON string: {ex.Message}");
                Console.WriteLine($"Error Code: {ex.Code}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred during import: {ex.Message}");
            }
        }
    }
}