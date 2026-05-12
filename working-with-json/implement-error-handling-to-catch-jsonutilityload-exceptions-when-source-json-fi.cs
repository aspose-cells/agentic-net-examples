using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonErrorHandling
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source JSON file (replace with your actual file path)
            string jsonFilePath = "input.json";

            // Create JSON load options (customize as needed)
            JsonLoadOptions loadOptions = new JsonLoadOptions
            {
                // Example option: start importing data from cell A1
                StartCell = "A1",
                // Keep the original JSON schema (optional)
                KeptSchema = true
            };

            try
            {
                // Attempt to load the JSON file into a Workbook.
                // This may throw an exception if the JSON is malformed.
                Workbook workbook = new Workbook(jsonFilePath, loadOptions);

                // If loading succeeds, you can work with the workbook here.
                Console.WriteLine("JSON file loaded successfully.");
                Console.WriteLine($"Number of worksheets: {workbook.Worksheets.Count}");

                // Example: save the workbook as an Excel file.
                workbook.Save("output.xlsx");
                Console.WriteLine("Workbook saved as output.xlsx.");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.FileCorrupted || ex.Code == ExceptionType.InvalidData)
            {
                // Specific handling for JSON parsing errors reported by Aspose.Cells.
                Console.WriteLine("Failed to load JSON file due to malformed content.");
                Console.WriteLine($"Error Code: {ex.Code}");
                Console.WriteLine($"Message: {ex.Message}");
            }
            catch (Exception ex)
            {
                // General fallback for any other unexpected errors.
                Console.WriteLine("An unexpected error occurred while loading the JSON file.");
                Console.WriteLine($"Message: {ex.Message}");
            }
        }
    }
}