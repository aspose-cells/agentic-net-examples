using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class JsonLoadWithErrorHandling
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            // Path to the source JSON file (may be malformed)
            string jsonPath = "malformed.json";

            // Verify that the JSON file exists to avoid FileNotFoundException
            if (!File.Exists(jsonPath))
            {
                Console.WriteLine($"JSON file not found: {jsonPath}");
                return;
            }

            // Create JSON load options (customize as needed)
            JsonLoadOptions loadOptions = new JsonLoadOptions
            {
                // Example option: start loading data from cell A1
                StartCell = "A1"
                // KeepSchema property does not exist in Aspose.Cells; omitted.
            };

            try
            {
                // Load the JSON file into a Workbook
                Workbook workbook = new Workbook(jsonPath, loadOptions);

                // Save the workbook to an Excel file
                workbook.Save("output.xlsx");
                Console.WriteLine("JSON loaded and saved successfully.");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.FileCorrupted ||
                                            ex.Code == ExceptionType.InvalidData ||
                                            ex.Code == ExceptionType.IO)
            {
                // Specific handling for JSON parsing related errors
                Console.WriteLine($"JSON loading failed: {ex.Message}");
                Console.WriteLine($"Exception Type: {ex.Code}");
            }
            catch (Exception ex)
            {
                // General fallback for any other unexpected errors
                Console.WriteLine($"An unexpected error occurred while loading JSON: {ex.Message}");
            }
        }
    }
}